from shutil import rmtree
from tempfile import mkdtemp


def prepare_mkdtemp(case):
    """
    Creates temporary directory and returns its name.
    """
    dirname = mkdtemp()
    case.addCleanup(rmtree, dirname)
    return dirname
