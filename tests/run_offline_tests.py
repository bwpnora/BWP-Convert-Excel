"""
Offline test runner for BWPConvertTTNVN.
Runs all static, ASCII safety, XML schema, and release packaging tests
that do not require an active Microsoft Excel COM installation.
Used by GitHub Actions CI/CD and developer pre-commit verification.
"""

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from tests.test_cell_processor import TestExcelAsciiSafety
from tests.test_currency_coordinator import TestCoreAsciiSafety
from tests.test_forms import TestFormsAsciiSafety
from tests.test_number_engine import TestNumberEngineStatic
from tests.test_release_packaging import (
    TestDocumentationAndMetadata,
    TestInstallerContract,
    TestReleasePackagingPipeline,
)
from tests.test_ribbon_package import (
    TestRibbonCallbacksSource,
    TestRibbonPackagingOpenXml,
    TestRibbonXmlContract,
)
from tests.test_udf_integration import TestUdfAsciiSafety
from tests.test_unicode_text import TestCoreTypes, TestUnicodeTextEngine


def main():
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    test_classes = [
        TestFormsAsciiSafety,
        TestCoreTypes,
        TestUnicodeTextEngine,
        TestDocumentationAndMetadata,
        TestInstallerContract,
        TestReleasePackagingPipeline,
        TestRibbonXmlContract,
        TestRibbonCallbacksSource,
        TestRibbonPackagingOpenXml,
        TestNumberEngineStatic,
        TestCoreAsciiSafety,
        TestExcelAsciiSafety,
        TestUdfAsciiSafety,
    ]

    for cls in test_classes:
        suite.addTests(loader.loadTestsFromTestCase(cls))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)


if __name__ == "__main__":
    main()
