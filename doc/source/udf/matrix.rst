.. Matrix-input functions

Matrix Functions
================

*All should work when taking range references. Quirky if typing arrays explicitly into the formula.*

.. function:: mTranspose(ByVal mtx As Variant) As Variant

   Wrapper around a call to the built-in ``MTRANSPOSE`` function that raises a customized RTE 13
   (type mismatch) if the argument is not a 2-D Array.


.. toctree::
   :maxdepth: 1