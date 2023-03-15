# 누수 사진 현황 만듭니다.

import win32com.client as win32

hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
# hwp.Open("C:/Users/k/Documents/ELIM/hwpComtrol/1.hwp","HWP","forceopen:true")


# 밑그림 그리기
stringArray = [
    "1) 누수(흔적) 사진현황",
    "T1,T2",
    "T3,T4",
    "T5,T6",
    "T7,T8",
    "T9,T10",
    "T11,T12",
    ""
    ]

for i in range(0, 8):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = stringArray[i]
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.Run("BreakPara")


# 표로 만들기
hwp.HAction.Run("MoveDocBegin")

hwp.HAction.Run("MoveDown")
hwp.HAction.Run("MoveSelDown")
hwp.HAction.Run("MoveSelDown")
hwp.HAction.Run("MoveSelDown")
hwp.HAction.Run("MoveSelDown")
hwp.HAction.Run("MoveSelDown")
hwp.HAction.Run("MoveSelDown")

hwp.HAction.GetDefault("TableStringToTable", hwp.HParameterSet.HTableStrToTbl.HSet)

hwp_HT = hwp.HParameterSet.HTableStrToTbl

hwp_HT.TableCreation.Rows = 6
hwp_HT.TableCreation.Cols = 2
hwp_HT.TableCreation.WidthType = 0
hwp_HT.TableCreation.HeightType = 0
hwp_HT.TableCreation.WidthValue = hwp.MiliToHwpUnit(148.0)
hwp_HT.TableCreation.HeightValue = hwp.MiliToHwpUnit(54.3)
hwp_HT.TableCreation.CreateItemArray("ColWidth", 2)

hwp_HT.TableCreation.CreateItemArray("RowHeight", 6)

# hwp_HT.TableCreation.ColWidth.Item[0] = hwp.MiliToHwpUnit(70.4)
# hwp_HT.TableCreation.ColWidth.Item[1] = hwp.MiliToHwpUnit(70.4)

# hwp_HT.TableCreation.RowHeight.Item[0] = hwp.MiliToHwpUnit(0.0)
# hwp_HT.TableCreation.RowHeight.Item[1] = hwp.MiliToHwpUnit(0.0)
# hwp_HT.TableCreation.RowHeight.Item[2] = hwp.MiliToHwpUnit(0.0)
# hwp_HT.TableCreation.RowHeight.Item[3] = hwp.MiliToHwpUnit(0.0)
# hwp_HT.TableCreation.RowHeight.Item[4] = hwp.MiliToHwpUnit(0.0)
# hwp_HT.TableCreation.RowHeight.Item[5] = hwp.MiliToHwpUnit(0.0)

hwp_HT.TableCreation.TableProperties.CellMarginLeft = hwp.MiliToHwpUnit(1.8)
hwp_HT.TableCreation.TableProperties.CellMarginRight = hwp.MiliToHwpUnit(1.8)
hwp_HT.TableCreation.TableProperties.CellMarginTop = hwp.MiliToHwpUnit(0.5)
hwp_HT.TableCreation.TableProperties.CellMarginBottom = hwp.MiliToHwpUnit(0.5)
hwp_HT.TableCreation.TableProperties.HorzRelTo = hwp.HorzRel("Column")
hwp_HT.TableCreation.TableProperties.VertRelTo = hwp.VertRel("Para")
hwp_HT.TableCreation.TableProperties.FlowWithText = 1
hwp_HT.TableCreation.TableProperties.TextWrap = hwp.TextWrapType("TopAndBottom")
hwp_HT.TableCreation.TableProperties.WidthRelTo = hwp.WidthRel("Absolute")
hwp_HT.TableCreation.TableProperties.HeightRelTo = hwp.HeightRel("Absolute")
hwp_HT.TableCreation.TableProperties.AllowOverlap = 0
hwp_HT.TableCreation.TableProperties.TreatAsChar = 0
hwp_HT.TableCreation.TableProperties.VertAlign = hwp.VAlign("Top")
hwp_HT.TableCreation.TableProperties.HorzAlign = hwp.HAlign("Justify")
hwp_HT.TableCreation.TableProperties.Width = 41954
hwp_HT.TableCreation.TableProperties.Height = 0
hwp_HT.TableCreation.TableProperties.TextFlow = hwp.TextFlowType("BothSides")
hwp_HT.TableCreation.TableProperties.OutsideMarginLeft = hwp.MiliToHwpUnit(1.0)
hwp_HT.TableCreation.TableProperties.OutsideMarginRight = hwp.MiliToHwpUnit(1.0)
hwp_HT.TableCreation.TableProperties.OutsideMarginTop = hwp.MiliToHwpUnit(1.0)
hwp_HT.TableCreation.TableProperties.OutsideMarginBottom = hwp.MiliToHwpUnit(1.0)
hwp_HT.TableCreation.TableProperties.HoldAnchorObj = 0
hwp_HT.AutoOrDefine = 1
hwp_HT.DelimiterType = hwp.Delimiter("SemiBreve")
hwp_HT.DelimiterEtc = ""
hwp_HT.UserDefine = ""

hwp.HAction.Execute("TableStringToTable", hwp.HParameterSet.HTableStrToTbl.HSet)



descriptionArray = [
    "지상4층 꿈누리터 천장",
    "지상4층 탁구실 천장",
    "지상3층 과학준비실 벽체",
    "지상1층 전기실 벽",
    "지상1층 펌프실 천장",
    "지상1층 펌프실 계단"
]


# 커서 이동
hwp.HAction.Run("MoveDocBegin")
hwp.HAction.Run("MoveDown")
hwp.HAction.Run("MoveDown")


for i in range(0, 6):

    # 칸 나누기
    hwp.HAction.GetDefault("TableSplitCell", hwp.HParameterSet.HTableSplitCell.HSet)
    hwp_TS = hwp.HParameterSet.HTableSplitCell
    hwp_TS.Rows = 0
    hwp_TS.Cols = 2
    hwp.HAction.Execute("TableSplitCell", hwp.HParameterSet.HTableSplitCell.HSet)

    # 칸 크기 변경
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp_SO = hwp.HParameterSet.HShapeObject

    hwp_SO.HSet.SetItem("ShapeType", 3)
    hwp_SO.HSet.SetItem("ShapeCellSize", 1)
    hwp_SO.ShapeTableCell.Width = hwp.MiliToHwpUnit(25.0)

    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)

    # 다음칸 편집
    hwp.HAction.Run("TableRightCellAppend")
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = descriptionArray[i]
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    
    if i%2==0:
        # 옆 칸으로
        hwp.HAction.Run("TableRightCellAppend")

    else:
        # 아래칸으로 
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("TableLeftCell")
        hwp.HAction.Run("MoveDown")


# === 여기까지 for문 반복 ===

j=1
k=1
for i in range(0,12):

    if (i//2)%2 == 0:
        hwp.HAction.Run("MoveDocEnd")
        hwp.InsertPicture("C:/Users/k/Documents/ELIM/hwpComtrol/img/"+str(k)+".jpg", Embedded=True)
        hwp.FindCtrl()
        hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)

        # 크기 변경
        hwp.HParameterSet.HShapeObject.Width = hwp.MiliToHwpUnit(70.0)
        hwp.HParameterSet.HShapeObject.Height = hwp.MiliToHwpUnit(65.0)
        hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)

        # 잘라내서 위치에 붙이기
        hwp.HAction.Run("Cut")
        hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
        hwp.HParameterSet.HFindReplace.FindString = "T"+str(i+1)
        hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

        hwp.HAction.GetDefault("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
        hwp.HAction.Execute("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
        k=k+1

    else :
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        option=hwp.HParameterSet.HFindReplace
        option.FindString = "T"+str(i+1)
        option.ReplaceString = "No."+str(j)
        option.IgnoreMessage = 1
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        j=j+1


file_root = "C:/Users/k/Documents/ELIM/Test"
new_filename = "test02.hwp"
new_file_path = file_root + "/" + new_filename
hwp.SaveAs(new_file_path)
hwp.Quit()