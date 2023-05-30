/* 
function：计算稿费合计
name：名字在第几列
value：统计的数值在第几列
*/
function price(dict, nameRow, valueRow, dispRow) {
    let arr = Sheets.Item(1).Range("A3:Z249").Value2
    let scores = {};
    for (let i = 0; i < arr.length; i++) {
        let name = arr[i][nameRow];
        let score = arr[i][valueRow];

        if (name != null) {
            if (name.includes("")) { // 判断是否有姓名
                let names = name.split(" "); //对当前单元格姓名进行分割
                let avgScore = Math.round(score / names.length * 100) / 100; //平均分
                for (let j = 0; j < names.length; j++) {
                    let n = names[j]; //单个单元格每个人的姓名
                    //console.log(n + "," + avgScore)
                    //console.log(n + "," +dict[n])
                    if (dict[n] == 0 || dict[n] == undefined) {
                        dict[n] = avgScore;
                        //console.log(n + "," +dict[n] + "," +avgScore)
                    } else {
                        dict[n] += avgScore;
                    }
                }
            }
        }
    }

    //	输出结果
    var i = 3;
    for (var key in dict) {
        console.log((i - 2) + "、" + key + "：" + dict[key])
            //console.log(key + "：" + dict[key]);
        Sheets.Item(2).Range("A" + i).Value2 = (i - 2); //序号
        Sheets.Item(2).Range("B" + i).Value2 = key; //姓名
        //Sheets.Item(2).Range("C" + i).Value2 = "=SUM(D" + i + ":H" + i + ")"; //
        Sheets.Item(2).Range("C" + i).Value2 = dict[key]; //总计
        i++;
    }
}

function main() {
    Columns(4).Insert();
    Sheets.Item(1).Range("D2").Value2 = "稿费";
    Columns(7).Insert()
    Columns(8).Insert()
    Sheets.Item(1).Range("G2").Value2 = "字数";
    Sheets.Item(1).Range("H2").Value2 = "稿费";
    Columns(10).Insert()
    Columns(10).Insert()
    Sheets.Item(1).Range("J2").Value2 = "张数";
    Sheets.Item(1).Range("K2").Value2 = "稿费";
    Columns(13).Insert()
    Sheets.Item(1).Range("M2").Value2 = "稿费";
    Columns(15).Insert()
    Sheets.Item(1).Range("O2").Value2 = "稿费";

    let rows = 14; //作品表初始化时的总行数
    for (let i = 3; i <= rows; i++) {
        Sheets.Item(1).Range("D" + i).Value2 = "=IF(COUNTA(C" + i + ")<>0,30,0)"; //编辑价格
        Sheets.Item(1).Range("H" + i).Value2 = "=G" + i + "/10"; //文字价格
        Sheets.Item(1).Range("K" + i).Value2 = "=J" + i + "*10"; //图片价格
        Sheets.Item(1).Range("M" + i).Value2 = "=IF(COUNTA(L" + i + ")<>0,80,0)"; //视频价格
        Sheets.Item(1).Range("O" + i).Value2 = "=IF(COUNTA(N" + i + ")<>0,50,0)"; //音频价格
    }

    // 合计行
    Sheets.Item(1).Range("A" + (rows + 1)).Value2 = "总计";
    Sheets.Item(1).Range("B" + (rows + 1)).Value2 = "=SUM(D" + (rows + 1) + ":P" + (rows + 1) + ")"; //总计
    //    Sheets.Item(1).Range("C" + (rows + 1)).Value2 = "=SUM(D" + (rows + 1) + ":N" +(rows+1)+")"; //总计
    Sheets.Item(1).Range("D" + (rows + 1)).Value2 = "=SUM(D1:D" + rows + ")"; //编辑稿费
    Sheets.Item(1).Range("H" + (rows + 1)).Value2 = "=SUM(H1:H" + rows + ")"; //文字稿费
    Sheets.Item(1).Range("K" + (rows + 1)).Value2 = "=SUM(K1:K" + rows + ")"; //图片稿费
    Sheets.Item(1).Range("M" + (rows + 1)).Value2 = "=SUM(M1:M" + rows + ")"; //视频稿费
    Sheets.Item(1).Range("O" + (rows + 1)).Value2 = "=SUM(O1:O" + rows + ")"; //音频稿费
    /// 设置表格标题颜色为红色
    Range("A" + (rows + 1) + ":P" + (rows + 1)).Select();
    (obj => {
        obj.Color = 255;
    })(Selection.Font)

    // 设置工作表的名称
    var shtName = "稿费代发表"
    var sht = Worksheets.Add(null, Sheets(Sheets.Count));
    sht.Name = shtName;

    // 设置表格标题
    var day = new Date();
    var month = day.getMonth() + 1;
    Sheets.Item(2).Range("A1").Value2 = "记者团" + (month - 1) + "月份积分表";
    /// 设置表格标题颜色为红色
    Range("A1:E1").Select();
    (obj => {
        obj.Color = 255;
    })(Selection.Font)

    //录入基本职责
    var jobName = ["序号", "姓名", "总计", "学号", "备注"];
    for (var letter = 65, i = 0; letter <= 77; letter++, i++) {
        Sheets.Item(2).Range(String.fromCharCode(letter) + "2").Value2 = jobName[i];
    }
    var dict = {}
    price(dict, 2, 3, "D");
    price(dict, 5, 7, "E");
    price(dict, 8, 10, "F");
    price(dict, 11, 12, "G");
    price(dict, 13, 14, "H");

    var endNum = Object.keys(dict).length + 3;
    Sheets.Item(2).Range("A" + endNum).Value2 = "总计";
    Sheets.Item(2).Range("B" + endNum).Value2 = "=SUM(C1:C" + (endNum - 1) + ")";
    Sheets.Item(2).Range("C" + endNum).Value2 = "=SUM(C1:C" + (endNum - 1) + ")";
    Sheets.Item(2).Range("E3:E" + (endNum - 1)).Value2 = "学生";
    /// 设置表格标题颜色为红色
    Range("A" + endNum + ":E" + endNum).Select();
    (obj => {
        obj.Color = 255;
    })(Selection.Font)
    //	console.log(dict[])
}
