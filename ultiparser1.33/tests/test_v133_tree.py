import unittest
from extract33 import SRC, def_block

class TestV133TreePopulate(unittest.TestCase):
    def test_depth_parameter_in_add_node(self):
        # Проверяем, что в add_node передаётся depth=0, а не обращение к несуществующему node.depth
        self.assertIn("def add_node(node: Node, parent_iid=\"\", depth=0):", SRC)
        self.assertNotIn("node.depth", SRC)

if __name__ == "__main__":
    unittest.main()
