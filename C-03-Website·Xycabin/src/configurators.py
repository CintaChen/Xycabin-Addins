"""Single source for configurator file names across the three projects."""

CONFIGURE = {
    'C-01-ExcelIterate': 'C1-Exceliterate_Configure.py',
    'C-02-WordUpdater': 'C2-Wordupdater_Configure.py',
    'C-03-Website·Xycabin': 'C3-XycabinSite_Configure.py',
}
BY_PRODUCT = {'excel': CONFIGURE['C-01-ExcelIterate'], 'word': CONFIGURE['C-02-WordUpdater']}


def name(folder):
    return CONFIGURE[folder]
