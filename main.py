"""نقطة تشغيل البرنامج / program entry point.

الاستخدام:
    python main.py                    -> يفتح الواجهة الرسومية
    python main.py chart.pptx         -> تحويل مباشر
    python main.py chart.pptx -o out.xlsx --tree
"""

import sys

from orgchart.cli import main

if __name__ == "__main__":
    sys.exit(main())
