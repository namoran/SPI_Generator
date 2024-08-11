# Copyright (c) 2010-2024 openpyxl

import pytest

@pytest.fixture
def Image():
    from ..image import Image
    return Image


class TestImage:

    @pytest.mark.pil_not_installed
    def test_import(self, Image, datadir):
        from ..image import _import_image
        datadir.chdir()
        with pytest.raises(ImportError):
            _import_image("plain.png")


    @pytest.mark.pil_required
    def test_ctor(self, Image, datadir):
        datadir.chdir()
        i = Image(img="plain.png")
        assert i.format == "PNG"
        assert i.width == 118
        assert i.height == 118
        assert i.anchor == "A1"


    @pytest.mark.pil_required
    def test_write_image(self, Image, datadir):
        datadir.chdir()
        i = Image("plain.png")
        with open("plain.png", "rb") as src:
            assert i._data() == src.read()


    @pytest.mark.pil_required
    def test_dont_close_pil(self, Image, datadir):
        datadir.chdir()
        from ..image import PILImage, Image
        obj = PILImage.open("plain.png")
        img = Image(obj)
        assert img.ref.fp is not None
        img._data()


    @pytest.mark.pil_required
    @pytest.mark.parametrize("filename, chars",
                             [
                                 ("plain.png", b'\x89PNG\r\n\x1a\n\x00\x00'),
                                 ("checkbox.emf", b"\x01\x00\x00\x00l\x00\x00\x00\x00\x00"),
                             ]
                             )
    def test_save(self, Image, datadir, filename, chars):
        datadir.chdir()
        img = Image(filename)
        assert img._data()[:10] == chars

    @pytest.mark.pil_required
    def test_save_with_alt_text(self, Image, datadir):
        datadir.chdir()
        img = Image("plain.png", "This alt text should not matter")
        assert img._data()[:10] == b'\x89PNG\r\n\x1a\n\x00\x00'

    @pytest.mark.pil_required
    def test_convert(self, Image, datadir):
        datadir.chdir()
        img = Image("plain.tif")
        assert img._data()[:10] == b'\x89PNG\r\n\x1a\n\x00\x00'


    @pytest.mark.pil_required
    def test_compare(self, Image, datadir):
        datadir.chdir()
        img1 = Image("plain.png")
        img2 = Image("plain.png")
        assert img1 == img2
