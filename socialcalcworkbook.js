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

    this.addNewSheet = function (name, str) {
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
        if (!this.currentSheetName) {
            this.currentSheetName = name;
        }
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
}