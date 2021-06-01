from unittest import TestCase, main
from local_window import make_file


class TestApp(TestCase):
    def test_generation(self):
        file_name = 'sample.xlsx'
        savedir = 'Out'
        int1, int2, int3, int4, int5, int6 = 10, 40, 10, 8, 8, 8
        self.assertTrue(make_file(file_name, savedir, int1, int2, int3, int4, int5, int6,
                                  {'label': 'Calibri'}, '10'))


if __name__ == '__main__':
    main()