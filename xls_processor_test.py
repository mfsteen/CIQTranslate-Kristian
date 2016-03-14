#!/usr/bin/env python3

import unittest
from xls_processor import Block, Blocks, CellCategory, Direction

class KnownBlocks(unittest.TestCase):
  known_blocks = (
    (Block(None, (0,0), (0,0), CellCategory.Label), (1,1)),
    (Block(None, (1,1), (1,14), CellCategory.Label), (1,14)),
    (Block(None, (2,1), (2,7), CellCategory.Input), (1, 7)),
    (Block(None, (5,0), (11,0), CellCategory.Label), (7, 1)),
    (Block(None, (5,1), (11,14), CellCategory.Intermediate), (7, 14))
  )

  def assertTupleEqual(self, t1, t2):
    self.assertEqual(len(t1), len(t2))
    for i in range(len(t1)):
      self.assertEqual(t1[i], t2[i])

  def assertBlockEqual(self, b1, b2):
    self.assertTupleEqual(b1.top_left, b2.top_left)
    self.assertTupleEqual(b1.dimensions(), b2.dimensions())

  def test_block_dimensions(self):
    for block, dim in self.known_blocks:
      self.assertTupleEqual(dim, block.dimensions())
      self.assertEqual(dim[0], block.row_count())
      self.assertEqual(dim[1], block.column_count())

  def test_next_block(self):
    blocks = Blocks()
    [blocks.add_block(block) for block, _ in self.known_blocks]

    block = self.known_blocks[4][0]
    nb = blocks.next_block(block, Direction.Top, (-1,block.column()), (0, block.column_count()), CellCategory.Label)
    self.assertBlockEqual(self.known_blocks[1][0], nb)

    nb = blocks.next_block(block, Direction.Left, (block.row(), -1), (block.row_count(), 1), CellCategory.Label)
    self.assertBlockEqual(self.known_blocks[3][0], nb)


if __name__ == '__main__':
  unittest.main()

