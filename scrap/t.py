import unittest

class DisjointError(Exception):
    pass
class ShorterError(Exception):
    pass

def relative_to(l1, l2):
    if len (l2) < len (l1):
        raise ShorterError("%s is shorter than %s" % (o, l2))
    for i1, i2 in zip(reversed(l1), reversed(l2)):
        if i1 != i2:
            raise DisjointError("%s is not relative to %s" % (l1, l2))
    else:
        return l2[:-len(l1)]

class TestRelativePath(unittest.TestCase):

    def test_shorter_is_error(self):
        self.assertRaises(ShorterError, relative_to, [1, 2], [1])

    def test_disjoint_is_error(self):
        self.assertRaises(DisjointError, relative_to, [1], [2])

    def test_equal_is_empty(self):
        expected = []
        answer = relative_to([1], [1])
        self.assertEqual(answer, expected)

    def test_true_relative(self):
        expected = [1, 2]
        answer = relative_to([3, 4], [1, 2, 3, 4])
        self.assertEqual(answer, expected)

if __name__ == '__main__':
    unittest.main()