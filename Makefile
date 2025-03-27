.PHONY: build build-windows

build:
	pyinstaller --name="SCGH XML Converter" \
		--windowed \
		--onefile \
		--exclude-module matplotlib \
		--exclude-module scipy \
		--exclude-module PIL \
		--exclude-module notebook \
		--exclude-module IPython \
		--exclude-module pytest \
		--exclude-module numpy.random.tests \
		--exclude-module numpy.core.tests \
		--exclude-module pandas.tests \
		--noupx \
		app.py

build-windows:
	docker run --rm -v "$(PWD):/src" cdrx/pyinstaller-windows:python3 "pip install -r requirements.txt && pyinstaller --name='SCGH XML Converter' --windowed --onefile --exclude-module matplotlib --exclude-module scipy --exclude-module PIL --exclude-module notebook --exclude-module IPython --exclude-module pytest --exclude-module numpy.random.tests --exclude-module numpy.core.tests --exclude-module pandas.tests --noupx app.py"