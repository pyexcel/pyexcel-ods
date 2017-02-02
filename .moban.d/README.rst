{%extends 'README.rst.jj2' %}

{%block description%}
**pyexcel-ods** is a tiny wrapper library to read, manipulate and write data in
ods format using python 2.6 and python 2.7. You are likely to use it with
`pyexcel <https://github.com/pyexcel/pyexcel>`_.
`pyexcel-ods3 <https://github.com/pyexcel/pyexcel-ods3>`_ is a sister library that
depends on ezodf and lxml. `pyexcel-odsr <https://github.com/pyexcel/pyexcel-odsr>`_
is the other sister library that has no external dependency but do ods reading only
{%endblock%}

{% block pagination_note%}
Special notice 30/01/2017: due to the constraints of the underlying 3rd party
library, it will read the whole file before returning the paginated data. So
at the end of day, the only benefit is less data returned from the reading
function. No major performance improvement will be seen.

With that said, please install `pyexcel-odsr <https://github.com/pyexcel/pyexcel-odsr>`_
and it gives better performance in pagination.
{%endblock%}

{%block extras %}
Credits
================================================================================

ODSReader is originally written by `Marco Conti <https://github.com/marcoconti83/read-ods-with-odfpy>`_
{%endblock%}
