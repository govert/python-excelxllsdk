from setuptools import setup, Extension, find_packages

setup(
    name='python-ExcelXLLSDK',
    use_vcs_version={'increment': '0.1'},
    author='Robert Thatcher',
    author_email='r.thatcher@cf-partners.com',
    ext_modules=[
        Extension('pelib.test._test_pelib', ['pelib/test/_test_pelib.c']),
    ],
    packages=find_packages(),
    package_data={
        '': ["*.xls"],
        'ExcelXLLSDK.test': ['test_builtins/*.xls']
    },
    entry_points={
        'console_scripts': [
            'pyxcel = pyxcel.__main__:main',
        ],
        'excel_addins': [
            'python_builtins = ExcelXLLSDK.builtins:DllMain',
            'python_unittest = ExcelXLLSDK.unittest:DllMain',
            'ExcelXLLSDK_test = ExcelXLLSDK.test.test_xll:DllMain',
            '_pyxcel = pyxcel.addin:DllMain',
        ]
    },
    license='LICENSE.txt',
    description='Create XLL add-ins with Python. Full access to the XLCALL API.',
    test_suite="nose.collector",
    install_requires=[
        'mock', # FIXME: prod deployment should not rely on mock
        'argh',
        'multimethod',
        'comtypes',
        'python-exceltools'
    ],
    setup_requires=[
        'hgtools',
        'nose',
        'mock',
    ],
)
