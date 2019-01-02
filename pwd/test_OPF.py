from unittest import TestCase
from doc_to_xml import OPF


class TestOPF(TestCase):
    opf = OPF('file.doc')

    def test_get_element_from_block(self):
        self.assertEqual('rishabh', self.opf.get_element_from_block(
            block=['email id: bhatnagarrishabh4@gmail.com', 'Name:- rishabh', 'age. 18'],
            identifier='name',
            split_by=':'
        ))
