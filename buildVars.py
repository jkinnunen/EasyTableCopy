# Build customizations
# Change this file instead of sconstruct or manifest files, whenever possible.

from site_scons.site_tools.NVDATool.typings import AddonInfo, BrailleTables, SymbolDictionaries
from site_scons.site_tools.NVDATool.utils import _

# Add-on information variables
addon_info = AddonInfo(
	# add-on Name/identifier, internal for NVDA
	addon_name="easyTableCopy",
	
	# Add-on summary/title, usually the user visible name of the add-on
	# Translators: Summary/title for this add-on
	addon_summary=_("Easy Table Copy"),
	
	# Add-on description
	# Translators: Long description to be shown for this add-on
	addon_description=_("""EasyTableCopy lets you copy tables, lists, and grids from the Web, desktop apps, and Windows Explorer into Word, Excel, or Outlook — with formatting intact. Mark specific rows or columns to copy, grab a single cell, flatten a tree view, or copy an entire folder listing, all with a keystroke."""),
	
	# version
	addon_version="2026.7.1",
	
	# Brief changelog for this version
	# Translators: what's new content for the add-on version
	addon_changelog=_("""
## Changed
- Improved reliability of quick/manual copy operations (marked rows/columns,
  cell copy, tree copy, list copy, SysListView32 lists) when the system
  clipboard is briefly held by another process. EasyTableCopy now retries a
  few times before giving up, instead of immediately reporting
  "Copy failed." This should eliminate occasional spurious copy failures
  caused by transient clipboard contention.

- Ensure 2026.2 compatibility.
"""),
	
	# Author(s)
	addon_author="Çağrı Doğan <cagrid@hotmail.com>",
	
	# URL for the add-on documentation support
	addon_url="https://github.com/Surveyor123/EasyTableCopy",
	
	# URL for the add-on repository where the source code can be found
	addon_sourceURL="https://github.com/Surveyor123/EasyTableCopy",
	
	# Documentation file name
	addon_docFileName="readme.html",
	
	# Minimum NVDA version supported
	addon_minimumNVDAVersion="2022.1.0",
	
	# Last NVDA version supported/tested
	addon_lastTestedNVDAVersion="2026.2.0",
	
	# Add-on update channel (None denotes stable releases)
	addon_updateChannel=None,
	
	# Add-on license
	addon_license="GPL-2.0",
	addon_licenseURL=None,
)

# Define the python files that are the sources of your add-on.
# We point to the specific directory where your code lives.
pythonSources: list[str] = ["addon/globalPlugins/EasyTableCopy/*.py"]

# Files that contain strings for translation. Usually your python sources
i18nSources: list[str] = pythonSources + ["buildVars.py"]

# Files that will be ignored when building the nvda-addon file
excludedFiles: list[str] = []

# Base language for the NVDA add-on
# Since your code strings (e.g. _("Table")) are in English, we keep this as "en".
baseLanguage: str = "en"

# Markdown extensions for add-on documentation
markdownExtensions: list[str] = []

# Custom braille translation tables
brailleTables: BrailleTables = {}

# Custom speech symbol dictionaries
symbolDictionaries: SymbolDictionaries = {}