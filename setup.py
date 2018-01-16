from setuptools import setup

setup(
    name='docxtool',
    version='0.1.0',
    description='A tool for manipulating docx outside of Word',
    author='Sergey Salnikov',
    author_email='serg@salnikov.ru',
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
    ],
    py_modules=['docxtool'],
    install_requires=[
        'cssutils',
        'python-docx',
    ],
    entry_points={
        'console_scripts': [
            'docxtool=docxtool:main',
        ],
    },
)
