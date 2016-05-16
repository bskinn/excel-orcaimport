.. Matrix-input functions

Matrix Functions
================

*All should work when taking range references; all require 2-D Array inputs.*

.. function:: mTranspose(ByVal mtx As Variant)

   *Returns:* 2-D Array

   *Worksheet function:* :ref:`mTranspose <udf-mtx-mTranspose>`

   Wrapper around a call to the built-in |xl-TRANSPOSE|_ function that raises a customized RTE 13
   (type mismatch) if the argument is not a 2-D Array.


.. toctree::
   :maxdepth: 1


