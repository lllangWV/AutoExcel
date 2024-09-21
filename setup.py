from setuptools import setup, find_packages

setup(
    name="section_writer",  
    version="0.1.0",  
    author="Logan Lang", 
    author_email="lllang@mix.wvu.edu", 
    description="This is a package for writing sections of papers.",  
    long_description=open("README.md").read(), 
    long_description_content_type="text/markdown",
    # url="https://github.com/yourusername/your_package_name",  # Replace with your package's URL
    packages=find_packages(),
    install_requires=[
        # Add your package dependencies here, e.g.:
        # "numpy>=1.21.0",
        # "pandas>=1.3.0",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",  # Replace with your license
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.9",  # Specify the minimum Python version required
)