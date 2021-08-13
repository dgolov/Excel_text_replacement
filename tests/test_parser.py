import unittest
from unittest.mock import Mock

from parsing import Parser


class TestParser(unittest.TestCase):
    def setUp(self) -> None:
        start_directory = '../tests'
        result_directory = '../tests'
        text_to_replace = 'to replace test'
        result_text = 'replaced'
        message_screen = Mock()
        self.parser = Parser(start_directory, result_directory, text_to_replace, result_text, message_screen)

    def test_run_no_such_directory(self):
        parser = Parser('', '', 'abc', 'cba')
        result = parser.run()
        self.assertEqual(result, 'Ошибка парсинга. Указанная папка не обнаружена!')

    def test_run_normal(self):
        result = self.parser.run()
        self.assertEqual(result, 'Парсинг завершен успешно!')

    def tearDown(self) -> None:
        pass
