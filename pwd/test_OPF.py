from unittest import TestCase
from doc_to_xml import OPF


class TestOPF(TestCase):
    opf = OPF('file.doc')

    def test_get_element_from_block(self):
        # Test when there is an element with given identifier.
        self.assertEqual(
            'rishabh',
            self.opf.get_element_from_block(
                block=['email id: bhatnagarrishabh4@gmail.com', 'Name:- rishabh', 'age. 18'],
                identifier='name',
                split_by=':'
            )
        )

        # Test when identifier doesn't exists in the given block.
        self.assertIsNone(
            self.opf.get_element_from_block(
                block=['email id: bhatnagarrishabh4@gmail.com', 'Name:- rishabh', 'age. 18'],
                identifier='address',
                split_by=':'
            )
        )
        return

    def test_position_insensitive_strip(self):
        # Test when some string of length greater than one is present in the original string.
        self.assertEqual(
            'rishabh',
            self.opf.position_insensitive_strip('.-+=rishabh:;!', weeds=';:.-+=!')
        )

        # Test when there are only weeds in the original string.
        self.assertEqual(
            '',
            self.opf.position_insensitive_strip('.-+=:;!', weeds=';+:.-=!')
        )

        # Test when there is only one element in the string to check if there is any off by one error in the function.
        self.assertEqual(
            '!',
            self.opf.position_insensitive_strip('..!..', weeds='.@.')
        )


