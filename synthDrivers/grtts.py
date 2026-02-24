# -*- coding: UTF-8 -*-
# synthDrivers/grtts.py - 广荣语音库 NVDA 驱动
# 广荣tts，NVDA语音合成器插件

import synthDriverHandler
from synthDriverHandler import SynthDriver, synthIndexReached, synthDoneSpeaking, VoiceInfo
from collections import OrderedDict
from logHandler import log

from synthDrivers import _grtts as _yongde

# 兼容性导入
try:
    from speech.commands import (
        BreakCommand, CharacterModeCommand, IndexCommand,
        LangChangeCommand, PitchCommand, RateCommand, VolumeCommand
    )
except ImportError:
    from speech import (
        BreakCommand, CharacterModeCommand, IndexCommand,
        LangChangeCommand, PitchCommand, RateCommand, VolumeCommand
    )

try:
    from autoSettingsUtils.driverSetting import BooleanDriverSetting, NumericDriverSetting, DriverSetting
except ImportError:
    from driverHandler import BooleanDriverSetting, NumericDriverSetting, DriverSetting


class SynthDriver(synthDriverHandler.SynthDriver):
    """广荣tts，NVDA语音合成器插件"""

    supportedSettings = (
        SynthDriver.VoiceSetting(),
        SynthDriver.RateSetting(),
        SynthDriver.VolumeSetting(),
        SynthDriver.PitchSetting(),
        NumericDriverSetting(
            "rateMultiplier",
            "语速倍率",
            minStep=1,
            normalStep=1,
            largeStep=2,
            defaultVal=1,
            minVal=1,
            maxVal=10,
            availableInSettingsRing=True
        ),
        BooleanDriverSetting(
            "dualVoice",
            "双音库模式",
            defaultVal=False,
            availableInSettingsRing=True
        ),
        DriverSetting(
            "englishVoice",
            "英文音库",
            availableInSettingsRing=True
        ),
        NumericDriverSetting(
            "englishRate",
            "英文语速",
            minStep=1,
            normalStep=5,
            largeStep=10,
            defaultVal=50,
            availableInSettingsRing=True
        ),
        NumericDriverSetting(
            "englishPitch",
            "英文音调",
            minStep=1,
            normalStep=5,
            largeStep=10,
            defaultVal=50,
            availableInSettingsRing=True
        ),
        NumericDriverSetting(
            "englishVolume",
            "英文音量",
            minStep=1,
            normalStep=5,
            largeStep=10,
            defaultVal=50,
            availableInSettingsRing=True
        ),
        NumericDriverSetting(
            "englishRateMultiplier",
            "英文语速倍率",
            minStep=1,
            normalStep=1,
            largeStep=2,
            defaultVal=1,
            minVal=1,
            maxVal=10,
            availableInSettingsRing=True
        ),
    )

    supportedCommands = {
        IndexCommand,
        CharacterModeCommand,
        BreakCommand,
        RateCommand,
        PitchCommand,
        VolumeCommand
    }

    supportedNotifications = {synthIndexReached, synthDoneSpeaking}

    description = '广荣tts'
    name = 'grtts'

    @classmethod
    def check(cls):
        """检查是否可用"""
        return _yongde.check()

    def __init__(self):
        """初始化"""
        try:
            _yongde.initialize(self._onIndexReached, self._onDoneSpeaking)
            log.info("grtts driver initialized")

            self._voice = _yongde.get_voice()
            self._rate = 50
            self._volume = 50
            self._pitch = _yongde.get_pitch()
            self._englishPitch = _yongde.get_english_pitch()

        except Exception as e:
            log.error(f"Failed to initialize grtts driver: {e}")
            raise

    def speak(self, speechSequence):
        """处理语音序列"""
        if not speechSequence:
            return

        chunks = []
        charmode = False
        pendingPitchReset = False

        def flush():
            nonlocal chunks, charmode, pendingPitchReset
            if not chunks:
                return
            if charmode:
                text = "".join(chunks)
                spelled = " ".join(ch for ch in text if not ch.isspace())
                if spelled:
                    _yongde.speak_sequence([("speak", spelled)])
            else:
                full_text = "".join(chunks).strip()
                if full_text:
                    _yongde.speak_sequence([("speak", full_text)])
            if pendingPitchReset:
                _yongde.speak_sequence([
                    ("pitch", self._pitch),
                    ("englishPitch", self._englishPitch),
                ])
                pendingPitchReset = False
            chunks = []

        for item in speechSequence:
            if isinstance(item, str):
                chunks.append(item)

            elif isinstance(item, IndexCommand):
                flush()
                _yongde.speak_sequence([("index", item.index)])

            elif isinstance(item, CharacterModeCommand):
                flush()
                charmode = item.state

            elif isinstance(item, BreakCommand):
                if item.time >= 500:
                    chunks.append("，")

            elif isinstance(item, RateCommand):
                flush()
                rate_val = max(0, min(item.newValue, 100))
                _yongde.set_rate(rate_val)

            elif isinstance(item, PitchCommand):
                flush()
                pitch_val = None
                english_val = None
                offset = getattr(item, "offset", None)
                has_offset = offset is not None and offset != 0
                # Prefer explicit offset if present (used for capitals).
                if has_offset:
                    try:
                        offset_val = int(offset)
                    except Exception:
                        offset_val = 0
                    pitch_val = self._pitch + offset_val
                    english_val = self._englishPitch + offset_val
                # NVDA may send PitchCommand with multiplier (newer) or newValue (older).
                elif hasattr(item, "multiplier"):
                    try:
                        multiplier = float(item.multiplier)
                    except Exception:
                        multiplier = 1.0
                    # Apply multiplier relative to each base pitch to preserve voice differences.
                    pitch_val = int(round(self._pitch * multiplier))
                    english_val = int(round(self._englishPitch * multiplier))
                elif hasattr(item, "newValue"):
                    pitch_val = int(item.newValue)
                    delta = pitch_val - self._pitch
                    english_val = self._englishPitch + delta
                else:
                    pitch_val = self._pitch
                    english_val = self._englishPitch
                pitch_val = max(0, min(pitch_val, 100))
                english_val = max(0, min(english_val, 100))
                _yongde.speak_sequence([("pitch", pitch_val), ("englishPitch", english_val)])
                pendingPitchReset = (
                    pitch_val != self._pitch or english_val != self._englishPitch
                )

            elif isinstance(item, VolumeCommand):
                flush()
                volume_val = max(0, min(item.newValue, 100))
                _yongde.set_volume(volume_val)

        flush()

    def pause(self, switch):
        """暂停/恢复"""
        _yongde.pause(switch)

    def cancel(self):
        """停止"""
        _yongde.stop()

    def terminate(self):
        """终止"""
        try:
            _yongde.terminate()
            log.info("grtts driver terminated")
        except Exception as e:
            log.error(f"Error terminating grtts driver: {e}")

    # 发音人设置
    def _getAvailableVoices(self):
        """获取可用语音"""
        voices = OrderedDict()
        voice_info = _yongde.get_available_voices()

        for voice_id, (name, internal_name) in voice_info.items():
            voices[voice_id] = VoiceInfo(voice_id, name, None)

        return voices

    def _get_voice(self):
        """获取当前语音"""
        return _yongde.get_voice()

    def _set_voice(self, voice_id):
        """设置语音"""
        if _yongde.set_voice(voice_id):
            self._voice = voice_id

    # 语速设置
    def _get_rate(self):
        """获取语速"""
        return _yongde.get_rate()

    def _set_rate(self, rate_percent):
        """设置语速"""
        _yongde.set_rate(rate_percent)

    # 音调设置
    def _get_pitch(self):
        """获取音调"""
        return _yongde.get_pitch()

    def _set_pitch(self, pitch_percent):
        """设置音调"""
        pitch_percent = max(0, min(pitch_percent, 100))
        self._pitch = pitch_percent
        _yongde.set_pitch(pitch_percent)

    # 音量设置
    def _get_volume(self):
        """获取音量"""
        return _yongde.get_volume()

    def _set_volume(self, volume_percent):
        """设置音量"""
        _yongde.set_volume(volume_percent)

    # 双音库模式设置
    def _get_dualVoice(self):
        """获取双音库开关"""
        return _yongde.get_dual_voice()

    def _set_dualVoice(self, value):
        """设置双音库开关"""
        _yongde.set_dual_voice(value)

    # 英文音库设置 - 直接复用主语音列表
    @property
    def availableEnglishVoices(self):
        """获取可用的英文音库 - 与主语音列表相同"""
        return self._getAvailableVoices()

    @property
    def availableEnglishvoices(self):
        """获取可用的英文音库 (小写v - NVDA兼容)"""
        return self._getAvailableVoices()

    def _get_englishVoice(self):
        """获取英文音库"""
        name = _yongde.get_english_voice_name()
        if name:
            return name
        # 如果没有设置，返回第一个可用的语音作为默认值
        voices = self._getAvailableVoices()
        if voices:
            first_voice = next(iter(voices.keys()))
            return first_voice
        return ""

    def _set_englishVoice(self, voice_id):
        """设置英文音库"""
        if voice_id:
            _yongde.set_english_voice_by_name(voice_id)

    # 英文语速设置
    def _get_englishRate(self):
        """获取英文语速"""
        return _yongde.get_english_rate()

    def _set_englishRate(self, rate_percent):
        """设置英文语速"""
        _yongde.set_english_rate(rate_percent)

    # 英文音调设置
    def _get_englishPitch(self):
        """获取英文音调"""
        return _yongde.get_english_pitch()

    def _set_englishPitch(self, pitch_percent):
        """设置英文音调"""
        pitch_percent = max(0, min(pitch_percent, 100))
        self._englishPitch = pitch_percent
        _yongde.set_english_pitch(pitch_percent)

    # 英文音量设置
    def _get_englishVolume(self):
        """获取英文音量"""
        return _yongde.get_english_volume()

    def _set_englishVolume(self, volume_percent):
        """设置英文音量"""
        _yongde.set_english_volume(volume_percent)

    # 语速倍率设置
    def _get_rateMultiplier(self):
        """获取语速倍率"""
        return _yongde.get_rate_multiplier()

    def _set_rateMultiplier(self, value):
        """设置语速倍率"""
        _yongde.set_rate_multiplier(value)

    # 英文语速倍率设置
    def _get_englishRateMultiplier(self):
        """获取英文语速倍率"""
        return _yongde.get_english_rate_multiplier()

    def _set_englishRateMultiplier(self, value):
        """设置英文语速倍率"""
        _yongde.set_english_rate_multiplier(value)

    # 回调
    def _onIndexReached(self, index):
        """索引到达"""
        if index is not None:
            synthIndexReached.notify(synth=self, index=index)
        else:
            synthDoneSpeaking.notify(synth=self)

    def _onDoneSpeaking(self):
        """朗读完成"""
        synthDoneSpeaking.notify(synth=self)
