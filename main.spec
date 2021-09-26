# -*- mode: python ; coding: utf-8 -*-

from kivy_deps import sdl2, glew, angle
import os
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.lang import Builder
from kivy.clock import Clock
from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime
from datetime import timedelta
from docx.shared import Pt
import sys
from os import path
from docx2pdf import convert
site_packages = next(p for p in sys.path if 'site-packages' in p)


block_cipher = None


a = Analysis(['main.py'],
             pathex=['C:\\prescription_v2'],
             binaries=[],
             datas=[(path.join(site_packages,"docx","templates"), "docx/templates")],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

a.datas += [('Code\main.kv','C:\\prescription_v2\main.kv', 'DATA')]

exe = EXE(pyz,
          a.scripts, 
          [],
          exclude_binaries=True,
          name='main',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None )

coll = COLLECT(exe,
               Tree('C:\\prescription_v2'),
               a.binaries,
               a.zipfiles,
               a.datas,
               *[Tree(p) for p in (sdl2.dep_bins + glew.dep_bins+ angle.dep_bins)],
               strip=False,
               upx=True,
               upx_exclude=[],
               name='main')
