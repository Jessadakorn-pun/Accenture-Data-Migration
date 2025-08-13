import os
from intern_lib  import (DeltaProcessorMock2, DeltaProcessorMock3, DeltaProcessorCutover)

# Point at your master control workbook:

# Delta Program Mock 2
# dp_mock2 = DeltaProcessorMock2(
#     master_path=r"C:\Users\USER\Desktop\Intern_WorkSpace\Master.xlsx"
# )

# Delta Program Mock 3
# dp_mock3 = DeltaProcessorMock3(
#     master_path=r"C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\Accenture-Data-Migration\Master3.xlsx"
# )

# dp_mock3.run()

# Delta Program Mock 4 (Cutover)
dp_mock4 = DeltaProcessorCutover(
    master_path=r"C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\2_Transform\MOCK4\00_DELTA\Master4.xlsx"
)

dp_mock4.run()