import os
from intern_lib  import DeltaProcessorMock2, DeltaProcessorMock3

# Point at your master control workbook:

# Delta Program Mock 2
# dp_mock2 = DeltaProcessorMock2(
#     master_path=r"C:\Users\USER\Desktop\Intern_WorkSpace\Master.xlsx"
# )

# Delta Program Mock 3
dp_mock3 = DeltaProcessorMock3(
    master_path=r"C:\Users\USER\Desktop\Intern_WorkSpace\Master.xlsx"
)

dp_mock3.run()
