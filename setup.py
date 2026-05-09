"""Setup configuration for RPA SharePoint Connector."""
from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as f:
    long_description = f.read()

setup(
    name="rpa-sharepoint-connector",
    version="1.0.0",
    description="Lean Python connector for RPA bots to work with SharePoint and OneDrive",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="RPA Team",
    packages=find_packages(),
    python_requires=">=3.7",
    install_requires=[
        "httpx>=0.25.0",
        "cryptography>=41.0.0",
    ],
    entry_points={
        "console_scripts": [
            "sharepoint-connector=rpa_sharepoint_connector.cli:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
)
