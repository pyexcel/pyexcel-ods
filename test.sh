#/bin/bash
pip freeze
python -m pytest --verbosity=3 --cov=pyexcel_ods --doctest-glob=*.rst
