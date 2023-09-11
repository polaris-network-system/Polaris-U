function findRow(sheet, val, col) {
    var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

    for (var i = 1; i < dat.length; i++) {
        if (dat[i][col - 1] == val) {
            return i + 1;
        }
    }
    return 0;
}

// Gmailの送信
function sendEmail(sheet, mailcol, consentcol, title, content) {
    // sheet : ユーザーDBのシートオブジェクト
    // mailcol : メールアドレスが記載されてる列
    // consentcol : メール配信同意(true or false)
    // title : メールのタイトル
    // contet : 本文、呼び出し元関数で生成する
    var emails = '';
    for (let i = 1; i < sheet.getLastRow() + 1; i++) {
        permission = sheet.getRange(i, consentcol - 1).getValue();
        if (
            (sheet.getRange(i, consentcol).getValue() == 'true' || sheet.getRange(i, consentcol).getValue() == 'TRUE' || sheet.getRange(i, consentcol).getValue() == true) &&
            (permission == 'Inside' || permission == 'Privilege' || permission == 'Admin' || permission == 'Advisor')
        ) {
            emails += sheet.getRange(i, mailcol).getValue() + ',';
        } else {
        }
    }
    var header = '<p class="navbar-item is-size-2 pt-1 pr-6" style="font-family: "Noto Sans JP", sans-serif; color: #004aad">Polaris-Uよりお知らせします</p>';
    var footer =
        '<br><br><br><br><strong>放送部活動支援プラットフォーム Polaris-U</strong><br><p>配信停止や配信の再開は<a href="https://script.google.com/a/oks.city-saitama.ed.jp/macros/s/AKfycbzmdIa50dzjEQYjV1Y4jKZS0PoEKeYFEZfpBHR8V0U/exec/settings_mail">こちら</a>からアクセスしてください<br>URLの末尾が「/exec/settings_email」になるようにしてアクセスすることでも可能です';
    var Draft = GmailApp.createDraft('', title, 'body', {
        name: 'Polaris-U',
        htmlBody: (header + content + footer).replaceAll('\n', '<br>'),
        bcc: emails,
    });
    Draft.send();
}

function findMultiRow(sheet, val, col) {
    var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
    var targetRows = [];
    var data = [];
    for (var i = 0; i < dat.length; i++) {
        if (dat[i][col - 1] == val) {
            targetRows.push(i + 1);
        }
    }
    targetRows = Array.from(new Set(targetRows));
    for (let i = 0; i < targetRows.length; i++) {
        // 検索にヒットしたレコードの取得
        let tmpdata = sheet.getRange(targetRows[i], 1, 1, sheet.getLastColumn()).getValues();
        data.push(tmpdata[0]);
    }
    return data;
}

function crossfindRow(sheet, key1, col1, key2, col2) {
    var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
    for (var i = 1; i < dat.length; i++) {
        if (dat[i][col1 - 1] == key1 && dat[i][col2 - 1] == key2) {
            return i + 1;
        }
    }
    return 0;
}

function graduation(sheet, key) {
    var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
    var result = [];
    for (var i = 0; i < dat.length; i++) {
        if (Number(dat[i][0]) < Number((key + 1) * 10000) && dat[i][5] == '部員') {
            result.push(i + 1);
        }
    }
    return result;
}

function findNearDataRow(sheet) {
    var result = [];
    var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

    //Dateオブジェクトからインスタンスを生成
    const today = new Date();
    for (var i = 0; i < dat.length; i++) {
        var dt = new Date(Number(dat[i][0]), Number(Number(dat[i][1]) - 1), Number(Number(dat[i][2]) + 1));
        if (dt < today) {
            result.push(i + 1);
        } else {
            break;
        }
    }
    var result_diff = 0;
    console.log(result);
    for (var j = 0; j < result.length; j++) {
        sheet.deleteRow(result[j] - result_diff);
        console.log(result[j]);
        result_diff += 1;
    }
}

function search(sheet, searchW) {
    let data = []; // 検索にヒットしたデータの格納先配列
    dat = sheet.getDataRange().getValues();
    if (searchW == null || searchW == '') {
        return dat; // 検索ワードがnullの場合は全件取得
    } else {
        ranges = sheet.createTextFinder(searchW).findAll(); // キーワードによる検索を実施

        let targetRows = []; // 検索にヒットしたレコード行の格納先
        data.push(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);

        // 検索にヒットしたRangeとレコード行を格納
        for (let i = 0; i < ranges.length; i++) {
            targetRows.push(ranges[i].getRow());
        }
        targetRows = Array.from(new Set(targetRows));
        for (let i = 0; i < targetRows.length; i++) {
            // 検索にヒットしたレコードの取得
            let tmpdata = sheet.getRange(targetRows[i], 1, 1, sheet.getLastColumn()).getValues();
            data.push(tmpdata[0]);
        }
        return data;
    }
}

function twoInt(number) {
    if (number.length < 2) {
        return '0' + number;
    } else {
        return number;
    }
}
function doGet(e) {
    // URLのexec/(またはdev/)以降を取得
    var page = e.pathInfo ? e.pathInfo : 'index';
    var settings_flag = '';

    // 設定変更用のURLを矯正
    if (page == 'settings_mail') {
        settings_flag = 'settings_email';
        page = 'index';
    } else {
    }
    // 該当するテンプレートを取得する
    var template = (() => {
        try {
            return HtmlService.createTemplateFromFile(page);
            //return HtmlService.createTemplateFromFile("templete");
        } catch (e) {
            return HtmlService.createTemplateFromFile('index');
        }
    })();

    var parameter = (() => {
        try {
            return e.parameter.page;
            //return HtmlService.createTemplateFromFile("templete");
        } catch (e) {
            return 'dummy';
        }
    })();

    var member_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部員登録情報');

    var LOGIN_USER = Session.getActiveUser().getEmail();
    try {
        var user_permission = member_db.getRange(findRow(member_db, LOGIN_USER, 5), 7).getValue();
        var user_name = member_db.getRange(findRow(member_db, LOGIN_USER, 5), 4).getValue();
    } catch {
        var user_permission = 'Outside';
        var user_name = '匿名';
    }

    // 設定変更
    if (settings_flag == 'settings_email') {
        var now_setting = member_db.getRange(findRow(member_db, LOGIN_USER, 5), 8).getValue();
        if (now_setting == 'true' || now_setting == true) {
            member_db.getRange(findRow(member_db, LOGIN_USER, 5), 8).setValue('false');
        } else if (now_setting == 'false' || now_setting == false) {
            member_db.getRange(findRow(member_db, LOGIN_USER, 5), 8).setValue('true');
        }
    }

    // htmlを返す
    template.page = parameter;
    template.user_name = user_name;
    template.user_permission = user_permission;
    template.url = ScriptApp.getService().getUrl(); // テンプレートにアプリのURLを渡す
    return template
        .evaluate() // テンプレートを評価してhtmlを返す
        .setTitle('Polaris-U') // タイトルをセット
        .addMetaTag('viewport', 'width=device-width,initial-scale=1'); // viewportを設定
}

function getData() {
    var LOGIN_USER = Session.getActiveUser().getEmail();
    var schedule_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部活日程');
    var member_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部員登録情報');
    var permission = member_db.getRange(findRow(member_db, LOGIN_USER, 5), 7).getValue();
    var item_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('機材情報');
    var absence_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('欠席連絡');
    var form_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('フォーム情報');
    switch (arguments[0]) {
        case 'index':
            findNearDataRow(schedule_db);
            schedule_db.getDataRange().sort({ column: 10, ascending: true });
            return schedule_db.getRange(2, 1, 1, schedule_db.getLastColumn()).getValues();

        case 'users':
            member_db.getRange(2, 1, member_db.getLastRow() - 1, member_db.getLastColumn()).sort({ column: 2, ascending: true });
            return search(member_db, '部員');

        case 'schedule_list':
            try {
                findNearDataRow(schedule_db);
                scheduleData = schedule_db.getRange(1, 1, schedule_db.getLastRow(), schedule_db.getLastColumn()).getValues();
                return [member_db.getRange(findRow(member_db, LOGIN_USER, 5), 7).getValue(), scheduleData, ''];
            } catch (e) {
                var message = '活動予定がありません';
                return [member_db.getRange(findRow(member_db, LOGIN_USER, 5), 7).getValue(), undefined, message];
            }

        case 'schedule_detail':
            var schedule = schedule_db.getRange(arguments[1], 1, 1, schedule_db.getLastColumn()).getValues();
            var key_date = twoInt(schedule[0][0]) + '-' + twoInt(schedule[0][1]) + '-' + twoInt(schedule[0][2]);
            var key_activity = schedule[0][4];
            var absence_list = absence_db.getDataRange().getValues();
            var absence_result = [];
            for (i = 1; i < absence_list.length; i++) {
                var check = absence_list[i][4].indexOf(key_activity);
                if (String(absence_list[i][3]) == key_date && check != -1) {
                    absence_result.push(absence_list[i]);
                }
            }

            // 活動作成者および特権・管理・顧問のみ編集・出席確認などの管理機能が使用可(allow)
            if (member_db.getRange(findRow(member_db, LOGIN_USER, 5), 4).getValue() == String(schedule[0][8]) || permission == 'Privilege' || permission == 'Admin' || permission == 'Advisor') {
                var permission = 'allow';
            } else {
                var permission = 'reject';
            }
            // ["欠席者","活動情報","権限"]
            return [absence_result, schedule, permission, arguments[1]];

        case 'item_list_default':
            return item_db.getDataRange().getValues();

        case 'absence_list':
            return search(absence_db, arguments[1]);

        case 'form_list':
            return form_db.getDataRange().getValues();

        case 'form_public':
            return findMultiRow(form_db, '受付中', 8);

        case 'item_list_search':
            return search(item_db, arguments[1]);

        case 'item_inquery':
            judge = findRow(item_db, arguments[1], 1);
            if (judge == 0) {
                return ['bad'];
            } else {
                return ['ok', item_db.getRange(judge, 2).getValue(), item_db.getRange(judge, 4).getValue(), item_db.getRange(judge, 5).getValue(), item_db.getRange(judge, 6).getValue()];
            }

        case 'absence_inquery':
            result = search(schedule_db, arguments[1]);
            if (result.length == 1) {
                return ['bad'];
            } else {
                return ['ok', result];
            }

        case 'absence_edit_inquery':
            result = absence_db.getRange(arguments[1], 1, 1, 6).getValues();
            if (result[0][0] == '') {
                return ['bad'];
            } else if (member_db.getRange(findRow(member_db, LOGIN_USER, 5), 1).getValue() != result[0][0]) {
                return ['notallow'];
            } else {
                return ['ok', result];
            }

        case 'form_inquery':
            result = form_db.getRange(arguments[1], 1, 1, 9).getValues();
            if (result[0][0] == '') {
                return ['bad'];
            } else {
                return ['ok', result];
            }

        case 'member_new_inquery':
            // コード認証
            judge = findMultiRow(member_db, arguments[1], 1);
            if (permission == 'Privilege' || permission == 'Admin' || permission == 'Advisor') {
                return ['bad'];
            } else if (judge.length != 0) {
                return ['already'];
            } else {
                // 学校アカウントのメールアドレス直書き
                return ['ok', arguments[1] + '@oks.city-saitama.ed.jp'];
            }

        case 'member_list_search':
            member_db.getRange(2, 1, member_db.getLastRow() - 1, member_db.getLastColumn() - 1).sort({ column: 2, ascending: true });
            return search(member_db, arguments[1]);

        case 'member_edit_inquery':
            // ["API名",学籍番号]
            result_row = findRow(member_db, arguments[1], 1);
            if (result_row == 0) {
                return ['bad'];
            } else {
                result = member_db.getRange(result_row, 1, 1, member_db.getLastColumn()).getValues();
                return ['ok', result[0]];
            }
    }
}

function sendData() {
    var LOGIN_USER = Session.getActiveUser().getEmail();
    var schedule_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部活日程');
    var member_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('部員登録情報');
    var item_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('機材情報');
    var absence_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('欠席連絡');
    var form_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('フォーム情報');
    findNearDataRow(schedule_db);

    switch (arguments[0]) {
        case 'schedule_new':
            var WeekChars = ['日', '月', '火', '水', '木', '金', '土'];
            var date = String(arguments[1]);
            var date_array = date.split('-');
            var hold_day = new Date(Number(date_array[0]), Number(date_array[1]) - 1, date_array[2].split('T')[0]);
            var day_youbi = WeekChars[hold_day.getDay()];
            schedule_db.appendRow([
                date_array[0],
                date_array[1],
                date_array[2].split('T')[0],
                day_youbi,
                String(date_array[2].split('T')[1]),
                arguments[2],
                arguments[3],
                arguments[4],
                member_db.getRange(findRow(member_db, LOGIN_USER, 5), 4).getValue(),
                arguments[1],
                schedule_db.getLastRow() + 1,
            ]);
            schedule_db.getRange(schedule_db.getLastRow(), 10).setNumberFormat('yyyy"-"mm"-"dd hh":"mm');
            schedule_db.getRange(schedule_db.getLastRow(), 1).setNumberFormat('@');
            schedule_db.getRange(schedule_db.getLastRow(), 2).setNumberFormat('@');
            schedule_db.getRange(schedule_db.getLastRow(), 3).setNumberFormat('@');
            schedule_db.getRange(schedule_db.getLastRow(), 5).setNumberFormat('@');
            schedule_db.getRange(schedule_db.getLastRow(), 10).setNumberFormat('@');

            // メール配信関係
            var title = '新規部活日程が追加されました';
            var signature = member_db.getRange(findRow(member_db, LOGIN_USER, 5), 4).getValue() + 'さんによって新規部活動予定が作成されました\n\n';
            var content = '活動内容：' + arguments[2];
            content += '\n活動日時:' + date_array[0] + '/' + date_array[1] + '/' + date_array[2].split('T')[0] + '(' + day_youbi + ') ' + date_array[2].split('T')[1];
            content += '\n活動場所：' + arguments[3];
            content += '\n備考　　：' + arguments[4];
            var url =
                '\n<a href="https://www.google.com/calendar/render?action=TEMPLATE&text=' +
                arguments[2] +
                '&dates=' +
                date_array[0] +
                date_array[1] +
                date_array[2].split(':')[0] +
                date_array[2].split(':')[1] +
                '00/' +
                date_array[0] +
                date_array[1] +
                date_array[2].split(':')[0].split('T')[0] +
                'T' +
                (Number(date_array[2].split(':')[0].split('T')[1]) + 1) +
                date_array[2].split(':')[1] +
                '00&details=' +
                arguments[4] +
                '">カレンダーに追加</a>\n\n';
            sendEmail(member_db, 5, 8, title, signature + content + url);
            return [arguments[2], String(arguments[1])];

        case 'schedule_update':
            var row = arguments[1];
            var WeekChars = ['日', '月', '火', '水', '木', '金', '土'];
            var date = String(arguments[2]);
            var date_array = date.split('-');
            var hold_day = new Date(Number(date_array[0]), Number(date_array[1]) - 1, date_array[2].split('T')[0]);
            var day_youbi = WeekChars[hold_day.getDay()];
            schedule_db
                .getRange(row, 1, 1, 8)
                .setValues([[date_array[0], date_array[1], date_array[2].split('T')[0], day_youbi, String(date_array[2].split('T')[1]), arguments[3], arguments[4], arguments[5]]]);
            schedule_db.getRange(row, 10, 1, 1).setValue(arguments[2]);
            schedule_db.getRange(row, 10).setNumberFormat('yyyy"-"mm"-"dd hh":"mm');
            schedule_db.getRange(row, 1).setNumberFormat('@');
            schedule_db.getRange(row, 2).setNumberFormat('@');
            schedule_db.getRange(row, 3).setNumberFormat('@');
            schedule_db.getRange(row, 5).setNumberFormat('@');
            schedule_db.getRange(row, 10).setNumberFormat('@');
            return [arguments[2], String(arguments[1])];

        case 'schedule_delete':
            var row = arguments[1];
            schedule_db.deleteRow(row);
            return;

        case 'item_new':
            item_db.appendRow([arguments[1], arguments[2], arguments[3], '健康', arguments[4], arguments[5]]);
            item_db.getRange(item_db.getLastRow(), 3).setNumberFormat('@');
            return [arguments[1], arguments[2]];

        case 'item_update':
            rownumber = findRow(item_db, arguments[1], 1);
            if (rownumber == 0) {
                return ['失敗しました', '該当の機材IDが見つかりませんでした'];
            } else {
                item_db.getRange(rownumber, 4, 1, 3).setValues([[String(arguments[3]), String(arguments[4]), String(arguments[5])]]);
            }
            return [arguments[1], arguments[2]];

        case 'item_update':
            rownumber = findRow(item_db, arguments[1], 1);
            if (rownumber == 0) {
                return ['失敗しました', '該当の機材IDが見つかりませんでした'];
            } else {
                item_db.getRange(rownumber, 4, 1, 3).setValues([[String(arguments[3]), String(arguments[4]), String(arguments[5])]]);
            }
            return [arguments[1], arguments[2]];

        case 'absence_new':
            //
            absence_db.appendRow([
                member_db.getRange(findRow(member_db, LOGIN_USER, 5), 1).getValue(),
                member_db.getRange(findRow(member_db, LOGIN_USER, 5), 4).getValue(),
                arguments[3],
                arguments[1],
                arguments[2],
                arguments[4],
                absence_db.getLastRow() + 1,
            ]);
            absence_db.getRange(absence_db.getLastRow(), 4).setNumberFormat('@');
            absence_db.getRange(absence_db.getLastRow(), 1).setNumberFormat('@');
            absence_db.getRange(absence_db.getLastRow(), 7).setNumberFormat('@');
            return [member_db.getRange(findRow(member_db, LOGIN_USER, 5), 4).getValue(), arguments[1], arguments[2]];

        case 'absence_new_slist':
            //[aid,atype,acontent,学籍番号(任意)]
            if (arguments[4]) {
                studentid = arguments[4];
                studentname = member_db.getRange(findRow(member_db, studentid, 1), 4).getValue();
            } else {
                studentid = member_db.getRange(findRow(member_db, LOGIN_USER, 5), 1).getValue();
                studentname = member_db.getRange(findRow(member_db, LOGIN_USER, 5), 4).getValue();
            }
            absence_db.appendRow([
                studentid,
                studentname,
                arguments[2],
                schedule_db.getRange(arguments[1], 10).getValue().split(' ')[0],
                schedule_db.getRange(arguments[1], 10).getValue().split(' ')[1] + ' ' + schedule_db.getRange(arguments[1], 6).getValue(),
                arguments[3],
                absence_db.getLastRow() + 1,
            ]);
            absence_db.getRange(absence_db.getLastRow(), 4).setNumberFormat('@');
            absence_db.getRange(absence_db.getLastRow(), 1).setNumberFormat('@');
            absence_db.getRange(absence_db.getLastRow(), 7).setNumberFormat('@');
            return [member_db.getRange(findRow(member_db, LOGIN_USER, 5), 4).getValue(), arguments[1], arguments[2]];

        case 'absence_delete':
            absence_db.deleteRow(arguments[1]);
            return [];

        case 'form_new':
            form_db.appendRow([
                arguments[1].split('-')[1],
                arguments[1].split('-')[2].split('T')[0],
                arguments[1].split('T')[1],
                arguments[2],
                arguments[3],
                arguments[4],
                arguments[5],
                '受付中',
                form_db.getLastRow() + 1,
            ]);
            form_db.getRange(form_db.getLastRow(), 1).setNumberFormat('@');
            form_db.getRange(form_db.getLastRow(), 2).setNumberFormat('@');
            form_db.getRange(form_db.getLastRow(), 3).setNumberFormat('@');
            return [arguments[2], arguments[3]];

        case 'form_update':
            form_db.getRange(arguments[1], 1, 1, 8).setValues([[arguments[2], arguments[3], arguments[4], arguments[5], arguments[6], arguments[7], arguments[8], arguments[9]]]);
            form_db.getRange(absence_db.getLastRow(), 1).setNumberFormat('@');
            form_db.getRange(absence_db.getLastRow(), 2).setNumberFormat('@');
            form_db.getRange(absence_db.getLastRow(), 3).setNumberFormat('@');
            return [];

        case 'form_delete':
            form_db.deleteRow(arguments[1]);
            return [];

        case 'member_new':
            //["API名","固有学籍番号","学籍暗号","部署","名前","メールアドレス"]
            member_db.appendRow([arguments[1], arguments[2], arguments[3], arguments[4], arguments[5], '部員', 'Inside']);
            member_db.getRange(form_db.getLastRow(), 1).setNumberFormat('@');
            member_db.getRange(form_db.getLastRow(), 2).setNumberFormat('@');
            return [arguments[1], arguments[4]];

        case 'member_update':
            member_db.getRange(findRow(member_db, arguments[1], 1), 1, 1, 7).setValues([[arguments[1], arguments[2], arguments[3], arguments[4], arguments[5], arguments[6], arguments[7]]]);
            member_db.getRange(member_db.getLastRow(), 1).setNumberFormat('@');
            member_db.getRange(member_db.getLastRow(), 2).setNumberFormat('@');
            return [arguments[1], arguments[4]];

        case 'member_upgrade':
            member_db.getRange(findRow(member_db, arguments[1], 1), 2).setValue(arguments[2]);
            member_db.getRange(member_db.getLastRow(), 2).setNumberFormat('@');
            return [];

        case 'member_graduetion':
            target = graduation(member_db, arguments[1]);
            for (var i = 0; i < target.length; i++) {
                member_db.getRange(target[i], 6).setValue('引退');
                member_db.getRange(target[i], 8).setValue('false');
                var tmp_class_number = member_db.getRange(target[i], 2).getValue();
                member_db.getRange(target[i], 2).setValue(Number('9' + String(tmp_class_number)));
            }
            return;

        case 'member_delete':
            member_db.deleteRow(arguments[1]);
            return [];

        case 'mail_send':
            var signature = '\n\nこのメールは' + member_db.getRange(findRow(member_db, LOGIN_USER, 5), 4).getValue() + 'さんによって作成されました\n\n';
            sendEmail(member_db, 5, 8, arguments[1], arguments[2] + signature);
            return ['ok'];
    }
}
