pip freeze
nosetests --with-cov --cover-package pyexcel_ods --cover-package tests --with-doctest --doctest-extension=.rst tests README.rst pyexcel_ods  && flake8 . --exclude=.moban.d 
