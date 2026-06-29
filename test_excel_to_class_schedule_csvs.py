import unittest

from excel_to_class_schedule_csvs import build_normalized_records


HEADERS = ['学 年', '学科・クラス', '月日', '時限', '変更内容', '科目(担当教員)']


def row(change_content: str, subject: str) -> dict[str, str]:
    return {
        '学 年': '3',
        '学科・クラス': 'IT',
        '月日': '2026/7/10',
        '時限': '1,2',
        '変更内容': change_content,
        '科目(担当教員)': subject,
    }


class ChangeContentFallbackTest(unittest.TestCase):
    def test_cancelled_class_is_the_before_subject(self) -> None:
        record = build_normalized_records(
            [row('休講', '基礎情報工学(担当者)')], HEADERS, default_year=2026
        )[0]
        self.assertEqual(record['before_subject'], '基礎情報工学(担当者)')
        self.assertEqual(record['after_subject'], '')
        self.assertEqual(record['note'], '休講')

    def test_makeup_and_changed_class_are_after_subjects(self) -> None:
        for change_content in ('補講', '変更'):
            with self.subTest(change_content=change_content):
                record = build_normalized_records(
                    [row(change_content, '数学IIIA(担当者)')], HEADERS, default_year=2026
                )[0]
                self.assertEqual(record['before_subject'], '')
                self.assertEqual(record['after_subject'], '数学IIIA(担当者)')
                self.assertEqual(record['note'], change_content)

    def test_explicit_before_after_columns_are_not_overwritten(self) -> None:
        headers = HEADERS + ['変更前', '変更後']
        source = row('変更', '補助情報') | {'変更前': '数学', '変更後': '物理'}
        record = build_normalized_records([source], headers, default_year=2026)[0]
        self.assertEqual(record['before_subject'], '数学')
        self.assertEqual(record['after_subject'], '物理')
        self.assertEqual(record['note'], '変更')


if __name__ == '__main__':
    unittest.main()
