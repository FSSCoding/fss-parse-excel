#!/usr/bin/env python3
"""
Setup script for Excel - Professional Spreadsheet Manipulation Toolkit
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read README for long description
readme_path = Path(__file__).parent / "README.md"
long_description = readme_path.read_text(encoding='utf-8') if readme_path.exists() else ""

# Read requirements
requirements_path = Path(__file__).parent / "requirements.txt"
requirements = []
if requirements_path.exists():
    requirements = requirements_path.read_text(encoding='utf-8').strip().split('\n')
    requirements = [req.strip() for req in requirements if req.strip() and not req.startswith('#')]

setup(
    name="fss-parse-excel",
    version="1.0.0",
    description="Professional Excel manipulation toolkit for CLI agents and automated workflows",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="System Architecture Team",
    author_email="development@example.com",
    url="https://github.com/FSSCoding/fss-parse-excel",
    
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    
    entry_points={
        'console_scripts': [
            'fss-parse-excel=excel_engine:main',
        ],
    },
    
    install_requires=requirements,
    
    extras_require={
        'dev': [
            'pytest>=6.0',
            'pytest-cov>=2.0',
            'black>=21.0',
            'flake8>=3.9',
            'mypy>=0.900',
        ],
        'docs': [
            'sphinx>=4.0',
            'sphinx-rtd-theme>=0.5',
        ],
    },
    
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Intended Audience :: System Administrators",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: System :: Systems Administration",
        "Topic :: Utilities",
    ],
    
    python_requires=">=3.8",
    
    keywords="excel spreadsheet xlsx csv automation cli agent",
    
    project_urls={
        "Bug Reports": "https://github.com/FSSCoding/fss-parse-excel/issues",
        "Source": "https://github.com/FSSCoding/fss-parse-excel",
        "Documentation": "https://github.com/FSSCoding/fss-parse-excel/blob/main/README.md",
    },
    
    include_package_data=True,
    zip_safe=False,
)