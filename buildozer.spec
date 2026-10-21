[app]
title = Ahmed Customs
package.name = ahmedcustoms
package.domain = org.customs
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 1.0.0
requirements = python3,kivy==2.3.0,openpyxl,et_xmlfile
orientation = portrait
fullscreen = 0
android.permissions = WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE
android.archs = arm64-v8a

[buildozer]
log_level = 2
warn_on_root = 1
