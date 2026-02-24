# -*- coding: UTF-8 -*-

import os

import addonHandler
import globalPluginHandler
import gui
from logHandler import log
import ui
import wx


addonHandler.initTranslation()


DONATE_MENU_LABEL = _("给广荣tts作者捐赠(&g)")
DONATE_DIALOG_TITLE = _("请用支付宝扫描中间二维码给广荣tts的作者捐赠")
COPY_BUTTON_LABEL = _("复制支付宝捐赠二维码到剪贴板")
MISSING_IMAGE_TEXT = _("未找到支付宝捐赠二维码图片。")
COPY_OK_TEXT = _("支付宝捐赠二维码已复制到剪贴板")
COPY_FAIL_TEXT = _("复制失败，请稍后重试")


class DonateDialog(wx.Dialog):
	def __init__(self, parent, imagePath):
		super().__init__(
			parent,
			title=DONATE_DIALOG_TITLE,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self._imagePath = imagePath
		self._bitmap = wx.NullBitmap
		self._copyButton = None
		self._buildUi()

	def _buildUi(self):
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		titleText = wx.StaticText(self, label=DONATE_DIALOG_TITLE)
		titleText.Wrap(560)
		mainSizer.Add(titleText, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 12)

		if os.path.isfile(self._imagePath):
			image = wx.Image(self._imagePath, wx.BITMAP_TYPE_ANY)
			if image.IsOk():
				maxWidth = 520
				maxHeight = 760
				width = image.GetWidth()
				height = image.GetHeight()
				if width > maxWidth or height > maxHeight:
					scale = min(float(maxWidth) / float(width), float(maxHeight) / float(height))
					image = image.Scale(
						int(width * scale),
						int(height * scale),
						wx.IMAGE_QUALITY_HIGH,
					)
				self._bitmap = wx.Bitmap(image)
				mainSizer.Add(
					wx.StaticBitmap(self, bitmap=self._bitmap),
					0,
					wx.ALL | wx.ALIGN_CENTER_HORIZONTAL,
					8,
				)
			else:
				mainSizer.Add(
					wx.StaticText(self, label=MISSING_IMAGE_TEXT),
					0,
					wx.ALL | wx.ALIGN_CENTER_HORIZONTAL,
					8,
				)
		else:
			mainSizer.Add(
				wx.StaticText(self, label=MISSING_IMAGE_TEXT),
				0,
				wx.ALL | wx.ALIGN_CENTER_HORIZONTAL,
				8,
			)

		buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
		self._copyButton = wx.Button(self, label=COPY_BUTTON_LABEL)
		self._copyButton.Enable(self._bitmap.IsOk())
		self._copyButton.Bind(wx.EVT_BUTTON, self._onCopyQrImage)
		buttonSizer.Add(self._copyButton, 0, wx.ALL, 8)

		closeButton = wx.Button(self, id=wx.ID_CLOSE)
		closeButton.Bind(wx.EVT_BUTTON, lambda evt: self.Close())
		buttonSizer.Add(closeButton, 0, wx.ALL, 8)

		mainSizer.Add(buttonSizer, 0, wx.ALIGN_CENTER_HORIZONTAL)
		self.SetSizerAndFit(mainSizer)
		self.CentreOnParent()
		self._copyButton.SetFocus()

	def _onCopyQrImage(self, evt):
		if not self._bitmap.IsOk():
			ui.message(COPY_FAIL_TEXT)
			return
		clipboard = wx.TheClipboard
		if not clipboard.Open():
			ui.message(COPY_FAIL_TEXT)
			return
		try:
			dataObject = wx.BitmapDataObject(self._bitmap)
			clipboard.SetData(dataObject)
			clipboard.Flush()
		finally:
			clipboard.Close()
		ui.message(COPY_OK_TEXT)


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	def __init__(self):
		super().__init__()
		self._menuItem = None
		self._sysTrayIcon = None
		try:
			self._sysTrayIcon = gui.mainFrame.sysTrayIcon
			menu = self._sysTrayIcon.menu
			insertPos = None
			for index in range(menu.GetMenuItemCount()):
				menuItem = menu.FindItemByPosition(index)
				if menuItem and menuItem.GetId() == wx.ID_EXIT:
					insertPos = index
					break
			if insertPos is None:
				self._menuItem = menu.Append(wx.ID_ANY, DONATE_MENU_LABEL)
			else:
				self._menuItem = menu.Insert(insertPos, wx.ID_ANY, DONATE_MENU_LABEL)
			self._sysTrayIcon.Bind(wx.EVT_MENU, self._onDonateMenu, self._menuItem)
		except Exception:
			log.error("Failed to register grtts donate menu item", exc_info=True)

	def terminate(self):
		try:
			if self._sysTrayIcon and self._menuItem:
				self._sysTrayIcon.Unbind(wx.EVT_MENU, handler=self._onDonateMenu, source=self._menuItem)
				menu = self._sysTrayIcon.menu
				menu.Remove(self._menuItem.GetId())
				self._menuItem.Destroy()
		except Exception:
			log.debugWarning("Failed to clean up grtts donate menu item", exc_info=True)
		self._menuItem = None
		self._sysTrayIcon = None
		super().terminate()

	def _onDonateMenu(self, evt):
		imagePath = os.path.join(os.path.dirname(__file__), "alipayDonateQr.jpg")
		gui.mainFrame.prePopup()
		try:
			dialog = DonateDialog(gui.mainFrame, imagePath)
			dialog.ShowModal()
			dialog.Destroy()
		except Exception:
			log.error("Failed to show grtts donate dialog", exc_info=True)
		finally:
			gui.mainFrame.postPopup()
