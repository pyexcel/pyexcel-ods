{%extends 'README.rst.jj2' %}

{%block result1%}
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [["row 1", "row 2", "row 3"]]}
{%endblock%}

{%block result2%}
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [[7, 8, 9], [10, 11, 12]]}
{%endblock%}

{%block description%}
**pyexcel-ods** is a tiny wrapper library to read, manipulate and write data in
ods format using python 2.6 and python 2.7. You are likely to use it with
`pyexcel <https://github.com/pyexcel/pyexcel>`_.
`pyexcel-ods3 <https://github.com/pyexcel/pyexcel-ods3>`_ is a sister library that
does the same thing but supports Python 3.3 and 3.4 and depends on lxml.
{%endblock%}

{%block extras %}
Credits
================================================================================

ODSReader is originally written by `Marco Conti <https://github.com/marcoconti83/read-ods-with-odfpy>`_
{%endblock%}
