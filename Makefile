PHONY: all

GITVERSION=`git describe --tags --always --abbrev=10 --dirty`

all: version
	pyinstaller crop2print.spec

version:
	echo "__version__ = '$(GITVERSION)'" > version.py
