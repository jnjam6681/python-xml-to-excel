.PHONY: build build-windows

install:
	pip3 install panda openpyxl lxml tk ttkthemes pyinstaller

build:
	pyinstaller --name="XML Converter" --windowed --onefile --noupx app.py

build-windows:
	docker run --rm -v "$(PWD):/src" cdrx/pyinstaller-windows:python3 "pip install -r requirements.txt && pyinstaller --name='XML Converter' --windowed --onefile --noupx app.py"