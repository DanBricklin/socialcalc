var SocialCalc;
if (!SocialCalc) {
    alert("Main SocialCalc code module needed");
    SocialCalc = {};
}
if (!SocialCalc.TableEditor) {
    alert("SocialCalc TableEditor code module needed");
}

SocialCalc.Workbook = function() {
    this.sheetInfoMap = {}; // name, sheet, context, editor, node
    this.currentSheetName = null;
    this.editor = null;

    this.getCurrentSheet = function () {
        return this.currentSheetName ? this.sheetInfoMap[this.currentSheetName].sheet : null;
    }

    this.addNewSheet = function (name, str) {
        var sheetInfo = this.sheetInfoMap[name];
        if (sheetInfo) {
            sheetInfo.sheet.ParseSheetSave(str);
            return sheetInfo.sheet;
        }
        var sheet = new SocialCalc.Sheet();
        SocialCalc.Formula.SheetCache.sheets[name] = {
            sheet: sheet,
            name: name
        };
        if (str) {
            sheet.ParseSheetSave(str);
        }
        var context = new SocialCalc.RenderContext(sheet);
        context.showGrid = true;
        context.showRCHeaders = true;
        var sheetInfo = {
            name: name,
            sheet: sheet,
            context: context,
        };
        this.sheetInfoMap[name] = sheetInfo;
        return sheet;
    }


    this.selectSheet = function (name, parentId, formulaId) {
        var sheetInfo = this.sheetInfoMap[name];
        if (!sheetInfo) {
            alert("no such sheet: " + name);
            return;
        }
        var editor = new SocialCalc.TableEditor(sheetInfo.context);
        var node = editor.CreateTableEditor(1800, 700);
        var inputbox = new SocialCalc.InputBox(document.getElementById(formulaId), editor);
        var parentNode = document.getElementById(parentId);
        for (child=parentNode.firstChild; child!=null; child=parentNode.firstChild) {
            parentNode.removeChild(child);
        }
        parentNode.appendChild(node);
        this.currentSheetName = name;
        this.editor = editor;
        editor.EditorScheduleSheetCommands("redisplay", true, false);
    }

    this.refresh = function () {
        this.editor.EditorScheduleSheetCommands("redisplay", true, false);
    }

    this.generateFile = function() {
        var json = {};
        for (var name in this.sheetInfoMap) {
            var sheetInfo = this.sheetInfoMap[name];
            var sheet = sheetInfo.sheet;
            var str = sheet.CreateSheetSave();
            json[name] = str;
        }
        return JSON.stringify(json);
    }

    this.addRemoteInfo = function (sheet, ref) {
        var range = this.editor.range;
        var info = {
            coord1: SocialCalc.crToCoord(range.left, range.top),
            coord2: SocialCalc.crToCoord(range.right, range.bottom),
            ref: ref,
        }
        sheet.remote.push(info);
        return info;
    }

    this.updateRemoteData = function (sheet, info, data) {
        var cr1 = SocialCalc.coordToCr(info.coord1);
        var cr2 = SocialCalc.coordToCr(info.coord2);
        var rcount = cr2.row - cr1.row + 1;
        if (rcount > data.length) {
            rcount = data.length;
        }
        var ccount = cr2.col - cr1.col + 1;
        for (var r = 0; r < rcount; ++r) {
            var row = data[r];
            for (var c = 0; c < row.length && c < ccount; ++c) {
                var coord = SocialCalc.crToCoord(c + cr1.col, r + cr1.row);
                var cell = sheet.GetAssuredCell(coord);
                cell.datavalue = row[c];
                if (!cell.datatype) {
                    if (r == 0) {
                        cell.datatype = "t";
                        cell.valuetype = "t";
                    } else {
                        cell.datatype = "v";
                        cell.valuetype = "n";
                    }
                }
            }
        }
    }
}