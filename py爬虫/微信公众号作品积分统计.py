function 统计学生所有积分() {
    // 设置工作表的名称
    var shtName = "积分表"
    var sht = Worksheets.Add(null, Sheets(Sheets.Count));
    sht.Name = shtName;

    // 设置表格标题
    var day = new Date();
    var month = day.getMonth() + 1;
    Sheets.Item(shtName).Range("A1").Value2 = "记者团" + (month - 1) + "月份积分表";
    /// 设置表格标题颜色为红色
    Range("A1").Select();
    (obj => {
        obj.Color = 255;
    })(Selection.Font)

    //录入基本职责
    var jobName = ["姓名", "编辑", "校对", "文字", "图片", "视频", "音频", "学生主编", "抖音", "活动", "总积分"];
    for (var letter = 65, i = 0; i < jobName.length; letter++, i++) {
        Sheets.Item(shtName).Range(String.fromCharCode(letter) + "2").Value2 = jobName[i];
    }

    //录入姓名
    let arr = Sheets.Item(1).Range("A3:I14").Value2
    let names = [];
    for (let j = 2; j <= 8; j++) {
        for (let i = 0; i < arr.length; i++) {
            let name = arr[i][j];
            Console.log("" + name)
            if (name != null) {
                if (name.includes("")) { // 判断是否有姓名
                    let n = name.split(" "); //对当前单元格姓名进行分割
                    for (let k = 0; k < n.length; k++) {
                        if (!(names.includes(n[k]))) {
                            names.push(n[k]);
                        }
                    }
                }
            }
        }
    }
    for (var i = 3, j = 0; names[j] != null; i++, j++) {
        Sheets.Item(shtName).Range("A" + i).Value2 = names[j];
    }

    //统计学生所有积分
    var sumLow = names.length + 3; //积分表所有的行数
    var cardNum = [4, 3, 6, 4, 5, 5, 5];
    var cardIndex = 0;
    for (var low = 66, integRow = 67; low != 72; low++, integRow++, cardIndex++) { //遍历积分表B2~H2
        for (var i = 3; i < sumLow; i++) { //遍历积分表A3~A100，得到所有姓名列
            var integration = 0;
            for (var j = 3; j < sumLow; j++) {
                var nameCell = Sheets.Item(1).Range(String.fromCharCode(integRow) + j).Value2; // 获取官微中的姓名
                var cell = Sheets.Item(shtName).Range("A" + i).Value2; //获取积分表的姓名  
                for (var a = 0; nameCell != null && nameCell[a] != null && nameCell[a + 1] != null; a++) {
                    var breakNum = 0;
                    for (var b = 0; cell != null && cell[b] != null; b++) {
                        if (nameCell[a] == cell[b] && nameCell[a + 1] == cell[b + 1]) {
                            integration++;
                            a++;
                            b++;
                        }
                    }
                }
            }
            // 每列的积分赋值
            Sheets.Item(shtName).Range(String.fromCharCode(low) + i).Value2 = integration * cardNum[cardIndex];
        }
    }

    //统计抖音的积分
    var low = 72;
    for (var i = 3; i < sumLow; i++) { //遍历积分表A3~A100，得到所有姓名列
        var integration = 0;
        for (integRow = 67; integRow != 69; integRow++) { //low遍历积分表B2~H2
            for (var j = 3; j < 100; j++) {
                var nameCell = Sheets.Item(2).Range(String.fromCharCode(integRow) + j).Value2; // 获取官微中的姓名
                var cell = Sheets.Item(shtName).Range("A" + i).Value2; //获取积分表的姓名  
                for (var a = 0; nameCell != null && nameCell[a] != null && nameCell[a + 1] != null; a++) {
                    var breakNum = 0;
                    for (var b = 0; cell != null && cell[b] != null; b++) {
                        if (nameCell[a] == cell[b] && nameCell[a + 1] == cell[b + 1]) {
                            integration++;
                            a++;
                            b++;
                        }
                    }
                }
            }
            // 每列的积分赋值
            Sheets.Item(shtName).Range(String.fromCharCode(low) + i).Value2 = integration * 4;
        }
    }

    //计算总积分
    for (var i = 3; i < sumLow; i++) {
        Sheets.Item(shtName).Range("k" + i).Value2 = "=sum(B" + i + ":J" + i + ")";
    }
}
