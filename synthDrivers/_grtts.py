# -*- coding: UTF-8 -*-
# synthDrivers/_grtts.py - 广荣语音库 TTS 底层实现 (DLL 版本)
# 广荣tts，NVDA语音合成器插件

import os
import ctypes
import threading
import queue
from os import path

import nvwave
from logHandler import log

# DLL 接口定义
_dll = None
_dll_path = None
_has_stream_chunk_ex = False

# 全局变量
initialized = False
player = None
speaking = False
on_index_reached = None
on_done_speaking = None
callback_queue = None
callback_thread = None
speak_queue = None
speak_lock = None
current_rate = 50
current_volume = 50
current_pitch = 50
current_english_pitch = 50
current_format = None
players_by_format = {}
speak_session_id = 0
current_speak_session = 0
session_lock = threading.Lock()
dll_lock = threading.RLock()
shutdown_event = threading.Event()
stop_call_lock = threading.Lock()
stop_call_thread = None
ACTION_SPEAK = "speak"
ACTION_INDEX = "index"
ACTION_RATE = "rate"
ACTION_VOLUME = "volume"
ACTION_PITCH = "pitch"
ACTION_ENGLISH_RATE = "englishRate"
ACTION_ENGLISH_VOLUME = "englishVolume"
ACTION_ENGLISH_PITCH = "englishPitch"
ACTION_RATE_MULTIPLIER = "rateMultiplier"
ACTION_ENGLISH_RATE_MULTIPLIER = "englishRateMultiplier"


def get_dll_path():
    """获取 DLL 所在目录"""
    return path.dirname(path.abspath(__file__))


def _load_dll():
    """加载 DLL"""
    global _dll, _dll_path, _has_stream_chunk_ex
    if _dll is not None:
        return True

    plugin_dir = get_dll_path()
    dll_file = path.join(plugin_dir, "grtts.dll")

    if not path.exists(dll_file):
        log.error(f"DLL not found: {dll_file}")
        return False

    try:
        _dll = ctypes.CDLL(dll_file)
        _dll_path = dll_file

        # 定义函数原型
        _dll.YongDe_Initialize.restype = ctypes.c_bool
        _dll.YongDe_Initialize.argtypes = []

        _dll.YongDe_Uninitialize.restype = None
        _dll.YongDe_Uninitialize.argtypes = []

        _dll.YongDe_GetVoiceCount.restype = ctypes.c_int
        _dll.YongDe_GetVoiceCount.argtypes = []

        _dll.YongDe_GetVoiceName.restype = ctypes.c_int
        _dll.YongDe_GetVoiceName.argtypes = [ctypes.c_int, ctypes.c_wchar_p, ctypes.c_int]

        _dll.YongDe_SetVoice.restype = ctypes.c_bool
        _dll.YongDe_SetVoice.argtypes = [ctypes.c_int]

        _dll.YongDe_GetVoice.restype = ctypes.c_int
        _dll.YongDe_GetVoice.argtypes = []

        _dll.YongDe_GetAudioFormat.restype = None
        _dll.YongDe_GetAudioFormat.argtypes = [
            ctypes.POINTER(ctypes.c_int),
            ctypes.POINTER(ctypes.c_int),
            ctypes.POINTER(ctypes.c_int)
        ]

        _dll.YongDe_TextToAudio.restype = ctypes.c_bool
        _dll.YongDe_TextToAudio.argtypes = [
            ctypes.c_wchar_p,
            ctypes.POINTER(ctypes.c_void_p),
            ctypes.POINTER(ctypes.c_int)
        ]

        _dll.YongDe_FreeAudio.restype = None
        _dll.YongDe_FreeAudio.argtypes = [ctypes.c_void_p]

        _dll.YongDe_SetRate.restype = None
        _dll.YongDe_SetRate.argtypes = [ctypes.c_int]

        _dll.YongDe_SetVolume.restype = None
        _dll.YongDe_SetVolume.argtypes = [ctypes.c_int]

        _dll.YongDe_SetPitch.restype = None
        _dll.YongDe_SetPitch.argtypes = [ctypes.c_int]

        _dll.YongDe_SetRateMultiplier.restype = None
        _dll.YongDe_SetRateMultiplier.argtypes = [ctypes.c_int]

        _dll.YongDe_GetRateMultiplier.restype = ctypes.c_int
        _dll.YongDe_GetRateMultiplier.argtypes = []

        # 双音库接口
        _dll.YongDe_SetDualVoice.restype = None
        _dll.YongDe_SetDualVoice.argtypes = [ctypes.c_bool]

        _dll.YongDe_GetDualVoice.restype = ctypes.c_bool
        _dll.YongDe_GetDualVoice.argtypes = []

        _dll.YongDe_SetEnglishVoice.restype = ctypes.c_bool
        _dll.YongDe_SetEnglishVoice.argtypes = [ctypes.c_int]

        _dll.YongDe_GetEnglishVoice.restype = ctypes.c_int
        _dll.YongDe_GetEnglishVoice.argtypes = []

        _dll.YongDe_SetEnglishRate.restype = None
        _dll.YongDe_SetEnglishRate.argtypes = [ctypes.c_int]

        _dll.YongDe_GetEnglishRate.restype = ctypes.c_int
        _dll.YongDe_GetEnglishRate.argtypes = []

        _dll.YongDe_SetEnglishPitch.restype = None
        _dll.YongDe_SetEnglishPitch.argtypes = [ctypes.c_int]

        _dll.YongDe_GetEnglishPitch.restype = ctypes.c_int
        _dll.YongDe_GetEnglishPitch.argtypes = []

        _dll.YongDe_SetEnglishVolume.restype = None
        _dll.YongDe_SetEnglishVolume.argtypes = [ctypes.c_int]

        _dll.YongDe_GetEnglishVolume.restype = ctypes.c_int
        _dll.YongDe_GetEnglishVolume.argtypes = []

        _dll.YongDe_SetEnglishRateMultiplier.restype = None
        _dll.YongDe_SetEnglishRateMultiplier.argtypes = [ctypes.c_int]

        _dll.YongDe_GetEnglishRateMultiplier.restype = ctypes.c_int
        _dll.YongDe_GetEnglishRateMultiplier.argtypes = []

        # 流式合成接口
        _dll.YongDe_StartStreamSynth.restype = ctypes.c_bool
        _dll.YongDe_StartStreamSynth.argtypes = [ctypes.c_wchar_p]

        _dll.YongDe_GetStreamChunk.restype = ctypes.c_bool
        _dll.YongDe_GetStreamChunk.argtypes = [
            ctypes.POINTER(ctypes.c_void_p),
            ctypes.POINTER(ctypes.c_int)
        ]

        try:
            _dll.YongDe_GetStreamChunkEx.restype = ctypes.c_bool
            _dll.YongDe_GetStreamChunkEx.argtypes = [
                ctypes.POINTER(ctypes.c_void_p),
                ctypes.POINTER(ctypes.c_int),
                ctypes.POINTER(ctypes.c_int),
                ctypes.POINTER(ctypes.c_int),
                ctypes.POINTER(ctypes.c_int)
            ]
            _has_stream_chunk_ex = True
        except AttributeError:
            _has_stream_chunk_ex = False

        _dll.YongDe_FreeStreamChunk.restype = None
        _dll.YongDe_FreeStreamChunk.argtypes = [ctypes.c_void_p]

        _dll.YongDe_IsStreamComplete.restype = ctypes.c_bool
        _dll.YongDe_IsStreamComplete.argtypes = []

        _dll.YongDe_StopStreamSynth.restype = None
        _dll.YongDe_StopStreamSynth.argtypes = []

        log.info(f"Loaded DLL: {dll_file}")
        return True

    except Exception as e:
        log.error(f"Failed to load DLL: {e}")
        _dll = None
        return False


def check():
    """检查是否可用"""
    if not _load_dll():
        return False
    plugin_dir = get_dll_path()

    # 检查广荣语音库(.vl)
    for f in os.listdir(plugin_dir):
        if f.endswith('.vl'):
            return True

    # 检查讯飞语音库(msc/res/tts/*.jet)
    xunfei_res = path.join(plugin_dir, 'msc', 'res', 'tts')
    if path.isdir(xunfei_res):
        common_jet = path.join(xunfei_res, 'common.jet')
        if path.exists(common_jet):
            # 检查是否有除common.jet外的其他音库
            for f in os.listdir(xunfei_res):
                if f.endswith('.jet') and f != 'common.jet':
                    return True

    # 检查 NaturalVoice 离线音库 (NaturalVoiceS/**/Tokens.xml)
    nv_dir = path.join(plugin_dir, 'NaturalVoiceS')
    if path.isdir(nv_dir):
        for root, _, files in os.walk(nv_dir):
            if 'Tokens.xml' in files:
                return True

    return False


def initialize(index_callback, done_callback):
    """初始化语音引擎"""
    global initialized, player, on_index_reached, on_done_speaking
    global callback_queue, callback_thread, speak_queue, speak_lock

    if initialized:
        return

    try:
        if not _load_dll():
            raise RuntimeError("Failed to load DLL")

        if not _dll.YongDe_Initialize():
            raise RuntimeError("DLL initialization failed")

        on_index_reached = index_callback
        on_done_speaking = done_callback

        player = create_player()
        callback_queue = queue.Queue()
        callback_thread = CallbackThread()
        callback_thread.daemon = True
        callback_thread.start()
        speak_queue = queue.Queue()
        speak_lock = threading.Lock()
        shutdown_event.clear()
        initialized = True

        log.info("grtts initialized successfully")

    except Exception as e:
        log.error(f"Init failed: {e}")
        raise


def terminate():
    """终止语音引擎"""
    global initialized, player, callback_queue, _dll

    if not initialized:
        return

    try:
        shutdown_event.set()
        stop()
        if callback_thread and callback_queue:
            callback_queue.put((None, None, None))
            callback_thread.join(timeout=5.0)

        if callback_thread and callback_thread.is_alive():
            log.warning("Callback thread still running; skip DLL uninitialize to avoid crash")
            initialized = False
            return

        _close_all_players()

        with dll_lock:
            if _dll:
                _dll.YongDe_Uninitialize()

        initialized = False
        log.info("grtts terminated")

    except Exception as e:
        log.error(f"Terminate error: {e}")


def speak(text):
    """朗读文本"""
    if not text or not text.strip():
        return False
    return speak_sequence([(ACTION_SPEAK, text.strip())])


def speak_sequence(actions):
    """Speak a sequence of actions."""
    if not initialized or not actions or not speak_queue or shutdown_event.is_set():
        return False

    queued = []
    for action in actions:
        if not action:
            continue
        action_type = action[0]
        if action_type == ACTION_SPEAK:
            text = action[1]
            if text and text.strip():
                queued.append((ACTION_SPEAK, text.strip()))
        elif action_type == ACTION_INDEX:
            index_value = action[1]
            if index_value is None:
                continue
            try:
                queued.append((ACTION_INDEX, int(index_value)))
            except (TypeError, ValueError):
                continue
        elif action_type in (
            ACTION_RATE,
            ACTION_VOLUME,
            ACTION_PITCH,
            ACTION_ENGLISH_RATE,
            ACTION_ENGLISH_VOLUME,
            ACTION_ENGLISH_PITCH,
            ACTION_RATE_MULTIPLIER,
            ACTION_ENGLISH_RATE_MULTIPLIER,
        ):
            try:
                queued.append((action_type, int(action[1])))
            except (TypeError, ValueError, IndexError):
                continue

    if not queued:
        return False

    try:
        with speak_lock:
            for item in queued:
                speak_queue.put(item)
            if speaking:
                return True
            started = _start_next_action_locked()
            return started
    except:
        return False


def _start_speak(text):
    """开始朗读"""
    global speaking, current_speak_session, speak_session_id

    try:
        if shutdown_event.is_set():
            return False
        with session_lock:
            speak_session_id += 1
            current_speak_session = speak_session_id
            session_id = current_speak_session
            speaking = True
        if callback_queue:
            callback_queue.put((_speak_async, (text, session_id), {}))
        return True
    except:
        _set_speaking(False)
        return False


def _start_next_action_locked():
    """Process the next queued action."""
    global speak_queue

    if not speak_queue:
        return False

    while not speak_queue.empty():
        action = speak_queue.get()
        if not action:
            continue
        action_type, value = action
        if action_type == ACTION_INDEX:
            if on_index_reached and callback_queue:
                callback_queue.put((on_index_reached, (value,), {}))
            continue
        if action_type == ACTION_RATE:
            set_rate(value)
            continue
        if action_type == ACTION_VOLUME:
            set_volume(value)
            continue
        if action_type == ACTION_PITCH:
            _set_pitch_runtime(value)
            continue
        if action_type == ACTION_RATE_MULTIPLIER:
            set_rate_multiplier(value)
            continue
        if action_type == ACTION_ENGLISH_RATE:
            set_english_rate(value)
            continue
        if action_type == ACTION_ENGLISH_VOLUME:
            set_english_volume(value)
            continue
        if action_type == ACTION_ENGLISH_PITCH:
            _set_english_pitch_runtime(value)
            continue
        if action_type == ACTION_ENGLISH_RATE_MULTIPLIER:
            set_english_rate_multiplier(value)
            continue
        if action_type == ACTION_SPEAK:
            return _start_speak(value)

    return False


def _is_current_session(session_id):
    if session_id is None:
        return True
    with session_lock:
        return session_id == current_speak_session


def _set_speaking(value, session_id=None):
    global speaking
    if session_id is None:
        speaking = value
        return True
    with session_lock:
        if session_id != current_speak_session:
            return False
        speaking = value
        return True


def _get_output_device():
    import config
    try:
        return config.conf["speech"]["outputDevice"]
    except Exception:
        pass
    try:
        return config.conf["audio"]["outputDevice"]
    except Exception:
        return "default"


def _remove_cached_player(target):
    if not target:
        return
    for fmt, cached in list(players_by_format.items()):
        if cached is target:
            players_by_format.pop(fmt, None)


def _close_player(target):
    if not target:
        return
    try:
        target.close()
    except Exception:
        pass
    _remove_cached_player(target)


def _close_all_players():
    global player, current_format
    players = list(players_by_format.values())
    if player and player not in players:
        players.append(player)
    players_by_format.clear()
    current_format = None
    player = None
    for p in players:
        try:
            p.close()
        except Exception:
            pass


def _speak_async(text, session_id=None):
    """异步朗读 - 使用流式合成"""
    global speaking, player, current_format

    try:
        if shutdown_event.is_set():
            return
        if not _is_current_session(session_id):
            return
        if not speaking or not _dll:
            return

        # 开始流式合成
        text_len = len(text) if text else 0
        with dll_lock:
            if shutdown_event.is_set() or not _dll:
                return
            started = _dll.YongDe_StartStreamSynth(text)
        if not started:
            log.warning(f"StartStreamSynth failed id={session_id} len={text_len}")
            if _set_speaking(False, session_id):
                if on_done_speaking:
                    callback_queue.put((on_done_speaking, (), {}))
            return

        # 循环获取音频块并播放
        while speaking and not shutdown_event.is_set():
            data_ptr = ctypes.c_void_p()
            data_len = ctypes.c_int()
            sample_rate = None
            bits_per_sample = None
            channels = None
            pcm_data = None
            got = False

            with dll_lock:
                if not _dll:
                    break
                if _dll.YongDe_IsStreamComplete():
                    break

                if _has_stream_chunk_ex:
                    sr = ctypes.c_int()
                    bps = ctypes.c_int()
                    ch = ctypes.c_int()
                    got = _dll.YongDe_GetStreamChunkEx(
                        ctypes.byref(data_ptr),
                        ctypes.byref(data_len),
                        ctypes.byref(sr),
                        ctypes.byref(bps),
                        ctypes.byref(ch)
                    )
                    if got:
                        sample_rate = sr.value
                        bits_per_sample = bps.value
                        channels = ch.value
                else:
                    got = _dll.YongDe_GetStreamChunk(ctypes.byref(data_ptr), ctypes.byref(data_len))

                if got and data_ptr.value and data_len.value > 0:
                    pcm_data = ctypes.string_at(data_ptr.value, data_len.value)
                    _dll.YongDe_FreeStreamChunk(data_ptr)

            if got and pcm_data:
                if _has_stream_chunk_ex and sample_rate and bits_per_sample and channels:
                    fmt = (sample_rate, bits_per_sample, channels)
                    if current_format != fmt:
                        old_player = player
                        new_player = create_player(
                            sample_rate=sample_rate,
                            bits_per_sample=bits_per_sample,
                            channels=channels
                        )
                        if old_player and old_player is not new_player:
                            try:
                                old_player.idle()
                            except Exception:
                                pass
                        if new_player:
                            player = new_player
                        else:
                            player = None

                if player and speaking and not shutdown_event.is_set():
                    try:
                        player.feed(pcm_data)
                    except FileNotFoundError:
                        _close_player(player)
                        player = create_player()
                        if player and speaking and not shutdown_event.is_set():
                            try:
                                player.feed(pcm_data)
                            except Exception:
                                pass
                    except Exception:
                        try:
                            player.idle()
                        except Exception:
                            pass
            else:
                import time
                time.sleep(0.001)

        # 等待播放完成
        if not _is_current_session(session_id):
            return
        if player and speaking:
            player.idle()

        if not _set_speaking(False, session_id):
            return
        next_started = _process_queue()
        if not next_started and on_done_speaking:
            callback_queue.put((on_done_speaking, (), {}))

    except Exception as e:
        if not shutdown_event.is_set():
            log.error(f"Speak error: {e}")
        if not _is_current_session(session_id):
            return
        if not _set_speaking(False, session_id):
            return
        _stop_stream_synth_safe(wait_timeout=0.05)
        if on_done_speaking:
            callback_queue.put((on_done_speaking, (), {}))


def _process_queue():
    """处理队列中的待朗读文本"""
    global speak_queue

    try:
        if shutdown_event.is_set():
            return False
        with speak_lock:
            started = _start_next_action_locked()
            return started
    except:
        pass
    return False


def _stop_stream_synth_safe(wait_timeout=0.1):
    """Best-effort stop that avoids blocking NVDA core indefinitely."""
    global stop_call_thread

    def _worker():
        try:
            with dll_lock:
                if _dll:
                    _dll.YongDe_StopStreamSynth()
        except Exception as e:
            if not shutdown_event.is_set():
                log.warning(f"StopStreamSynth error: {e}")

    worker = None
    with stop_call_lock:
        if stop_call_thread and stop_call_thread.is_alive():
            worker = stop_call_thread
        else:
            stop_call_thread = threading.Thread(
                target=_worker,
                name="YongDeStopStream",
                daemon=True
            )
            stop_call_thread.start()
            worker = stop_call_thread

    if wait_timeout and wait_timeout > 0:
        worker.join(wait_timeout)
        if worker.is_alive() and not shutdown_event.is_set():
            log.warning("YongDe_StopStreamSynth timed out; continue without blocking core")
            return False

    return True


def stop():
    """停止朗读"""
    global speaking, player

    if speak_queue:
        with speak_lock:
            while not speak_queue.empty():
                try:
                    speak_queue.get_nowait()
                except:
                    break

    _set_speaking(False)

    # 停止流式合成
    _stop_stream_synth_safe(wait_timeout=0.1)

    players = list(players_by_format.values())
    if player and player not in players:
        players.append(player)
    for p in players:
        try:
            p.stop()
        except Exception:
            pass

    _restore_pitch_baseline()


def pause(switch):
    """暂停/恢复朗读"""
    if player:
        player.pause(switch)


class CallbackThread(threading.Thread):
    """回调处理线程"""
    def run(self):
        while True:
            try:
                func, args, kwargs = callback_queue.get(timeout=1.0)
                if func is None:
                    break
                try:
                    if kwargs:
                        kwargs.pop("mustBeAsync", None)
                    func(*args, **(kwargs or {}))
                except:
                    pass
                callback_queue.task_done()
            except:
                continue


def create_player(sample_rate=None, bits_per_sample=None, channels=None):
    """创建音频播放器"""
    global current_format
    try:
        if sample_rate is None or bits_per_sample is None or channels is None:
            # 获取音频格式
            sr = ctypes.c_int()
            bps = ctypes.c_int()
            ch = ctypes.c_int()

            if _dll:
                _dll.YongDe_GetAudioFormat(
                    ctypes.byref(sr),
                    ctypes.byref(bps),
                    ctypes.byref(ch)
                )
            else:
                sr.value = 16000
                bps.value = 16
                ch.value = 1
            sample_rate = sr.value
            bits_per_sample = bps.value
            channels = ch.value

        fmt = (sample_rate, bits_per_sample, channels)
        cached = players_by_format.get(fmt)
        if cached:
            current_format = fmt
            return cached
        dev = _get_output_device()
        new_player = nvwave.WavePlayer(channels, sample_rate, bits_per_sample, outputDevice=dev)
        players_by_format[fmt] = new_player
        current_format = fmt
        return new_player
    except Exception as e:
        log.error(f"Failed to create player: {e}")
        return None


def get_available_voices():
    """获取可用语音列表"""
    result = {}

    if not _dll:
        return result

    count = _dll.YongDe_GetVoiceCount()
    for i in range(count):
        # 使用更大的缓冲区以支持带前缀的语音名称（如 "[讯飞]小燕"）
        buffer = ctypes.create_unicode_buffer(256)
        _dll.YongDe_GetVoiceName(i, buffer, 256)
        name = buffer.value
        result[name] = (name, name)

    return result


def get_voice():
    """获取当前语音"""
    if not _dll:
        return None

    idx = _dll.YongDe_GetVoice()
    count = _dll.YongDe_GetVoiceCount()

    if idx >= 0 and idx < count:
        buffer = ctypes.create_unicode_buffer(128)
        _dll.YongDe_GetVoiceName(idx, buffer, 128)
        return buffer.value

    return None


def set_voice(v):
    """设置当前语音"""
    global player
    if not _dll:
        return False

    count = _dll.YongDe_GetVoiceCount()
    for i in range(count):
        buffer = ctypes.create_unicode_buffer(128)
        _dll.YongDe_GetVoiceName(i, buffer, 128)
        if buffer.value == v:
            if _dll.YongDe_SetVoice(i):
                try:
                    stop()
                    if player:
                        _close_player(player)
                except Exception:
                    pass
                player = create_player()
                return True
            return False

    return False


def get_rate():
    """获取语速"""
    return current_rate


def set_rate(v):
    """设置语速"""
    global current_rate
    current_rate = max(0, min(100, v))
    if _dll:
        _dll.YongDe_SetRate(current_rate)
    return True


def get_volume():
    """获取音量"""
    return current_volume


def set_volume(v):
    """设置音量"""
    global current_volume
    current_volume = max(0, min(100, v))
    if _dll:
        _dll.YongDe_SetVolume(current_volume)
    return True


def get_pitch():
    """获取音高"""
    return current_pitch


def set_pitch(v):
    """设置音高"""
    global current_pitch
    current_pitch = max(0, min(100, v))
    _set_pitch_runtime(current_pitch)
    return True


def _set_pitch_runtime(v):
    if _dll:
        _dll.YongDe_SetPitch(max(0, min(100, v)))


def _set_english_pitch_runtime(v):
    if _dll:
        _dll.YongDe_SetEnglishPitch(max(0, min(100, v)))


def _restore_pitch_baseline():
    _set_pitch_runtime(current_pitch)
    _set_english_pitch_runtime(current_english_pitch)


# ============== 双音库接口 ==============

def get_dual_voice():
    """获取双音库开关状态"""
    if not _dll:
        return False
    return _dll.YongDe_GetDualVoice()


def set_dual_voice(enable):
    """设置双音库开关"""
    if _dll:
        _dll.YongDe_SetDualVoice(enable)
    return True


def get_english_voice():
    """获取英文音库索引"""
    if not _dll:
        return -1
    return _dll.YongDe_GetEnglishVoice()


def set_english_voice(index):
    """设置英文音库索引"""
    if not _dll:
        return False
    return _dll.YongDe_SetEnglishVoice(index)


def get_english_voice_name():
    """获取英文音库名称"""
    if not _dll:
        return None

    idx = _dll.YongDe_GetEnglishVoice()
    count = _dll.YongDe_GetVoiceCount()

    if idx >= 0 and idx < count:
        buffer = ctypes.create_unicode_buffer(256)
        _dll.YongDe_GetVoiceName(idx, buffer, 256)
        return buffer.value

    return None


def set_english_voice_by_name(name):
    """通过名称设置英文音库"""
    if not _dll:
        return False

    count = _dll.YongDe_GetVoiceCount()
    for i in range(count):
        buffer = ctypes.create_unicode_buffer(256)
        _dll.YongDe_GetVoiceName(i, buffer, 256)
        if buffer.value == name:
            return _dll.YongDe_SetEnglishVoice(i)

    return False


def get_english_rate():
    """获取英文语速"""
    if not _dll:
        return 50
    return _dll.YongDe_GetEnglishRate()


def set_english_rate(v):
    """设置英文语速"""
    if _dll:
        _dll.YongDe_SetEnglishRate(max(0, min(100, v)))
    return True


# ============== 倍速接口 ==============

def get_rate_multiplier():
    """获取语速倍率"""
    if not _dll:
        return 1
    return _dll.YongDe_GetRateMultiplier()


def set_rate_multiplier(v):
    """设置语速倍率 (1-10)"""
    if _dll:
        _dll.YongDe_SetRateMultiplier(max(1, min(10, v)))
    return True


# ============== 英文音调/音量接口 ==============

def get_english_pitch():
    """获取英文音调"""
    return current_english_pitch


def set_english_pitch(v):
    """设置英文音调"""
    global current_english_pitch
    current_english_pitch = max(0, min(100, v))
    _set_english_pitch_runtime(current_english_pitch)
    return True


def get_english_volume():
    """获取英文音量"""
    if not _dll:
        return 50
    return _dll.YongDe_GetEnglishVolume()


def set_english_volume(v):
    """设置英文音量"""
    if _dll:
        _dll.YongDe_SetEnglishVolume(max(0, min(100, v)))
    return True


# ============== 英文倍速接口 ==============

def get_english_rate_multiplier():
    """获取英文语速倍率"""
    if not _dll:
        return 1
    return _dll.YongDe_GetEnglishRateMultiplier()


def set_english_rate_multiplier(v):
    """设置英文语速倍率 (1-10)"""
    if _dll:
        _dll.YongDe_SetEnglishRateMultiplier(max(1, min(10, v)))
    return True
