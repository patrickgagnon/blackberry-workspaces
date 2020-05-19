from setuptools import setup

with open("README.md", "r") as readme_file:
    readme = readme_file.read()

setup(
    name="blackberry-workspaces",
    version="0.0.1",
    author="Patrick Gagnon",
    auther_email="plgagnon00@gmail.com",
    description="A package to interact with Blackberry Workspaces via their REST API",
    long_description=readme,
    url="https://github.com/patrickgagnon/blackberry-workspaces",
    packages=['workspaces'],
)
