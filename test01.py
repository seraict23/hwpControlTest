# pip install win32
import win32com.client as win32

hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
# hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
hwp.Open("C:/Users/k/Documents/ELIM/hwpComtrol/1.hwp","HWP","forceopen:true")



hwp.HAction.Run("MoveDocEnd")

hwp.HAction.Run("MoveSelUp")
hwp.HAction.Run("MoveSelUp")
hwp.HAction.Run("MoveSelUp")
hwp.HAction.Run("MoveSelUp")
hwp.HAction.Run("MoveSelUp")
hwp.HAction.Run("MoveSelUp")
hwp.HAction.Run("MoveSelUp")

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


for i in range(1,7):
    hwp.HAction.Run("MoveDocEnd")
    hwp.InsertPicture("C:/Users/k/Documents/ELIM/hwpComtrol/img/"+str(i)+".jpg", Embedded=True)
    hwp.FindCtrl()
    hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)

    # 크기 변경
    hwp.HParameterSet.HShapeObject.Width = hwp.MiliToHwpUnit(65.0)
    hwp.HParameterSet.HShapeObject.Height = hwp.MiliToHwpUnit(90.0)
    hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)

    # 잘라내서 위치에 붙이기
    hwp.HAction.Run("Cut")
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = str(i)
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

    hwp.HAction.GetDefault("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
    hwp.HAction.Execute("Paste", hwp.HParameterSet.HSelectionOpt.HSet)



file_root = "C:/Users/k/Documents/ELIM/hwpComtrol"
new_filename = "new1.hwp"
new_file_path = file_root + "/" + new_filename
hwp.SaveAs(new_file_path)
hwp.Quit()