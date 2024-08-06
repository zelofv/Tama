(() => {
    "use strict";
    GM_deleteValue("currentType");
    const oldFetch = fetch;
    // 拦截请求，根据请求判断渲染表格，总共50行数据
    unsafeWindow.fetch = (...args) => {
        const url = args[0];
        const uri = url.split("?")[0];
        const params = url.split("?")[1].split("&");
        const _seed = params[0].split("=")[1].replaceAll(" ", "");
        const _index = params[1].split("=")[1].replaceAll(" ", "");
        for (let i = 0; i < params.length; i++) {
            if (params[i].indexOf("=") + 1 === params[i].length) {
                return;
            }
        }

        const type = uri.substring(uri.lastIndexOf("/") + 1, uri.length);
        // const list = JSON.parse(GM_getValue("searchParams") || "{}");
        const map = JSON.parse(GM_getValue("searchParamsMap") || "{}");
        map[type] = map[type] || {};
        const list = map[type][_seed] || [];
        if (!list.includes(_index)) {
            list.unshift(_index);
            map[type][_seed] = list;
            // GM_setValue("searchParams", JSON.stringify(list));
            GM_setValue("searchParamsMap", JSON.stringify(map));
        }
        // 在这里清空表格格式
        setTimeout(() => {
            clearAllStyle(oldTbody, type);
        }, 0);
        return new Promise((resolve, reject) => {
            oldFetch(...args).then((res) => {
                const oldJson = res.json;
                res.json = function() {
                    return new Promise((resolve, reject) => {
                        oldJson.apply(this, arguments).then((result) => {
                            // console.log(result);
                            // result.hook = "success";
                            resolve(result);
                        });
                    });
                };
                resolve(res);
            }).then((res) => {
                GM_setValue("currentType", type);
                if (url?.includes("RandomMainAbility")) {
                    setTimeout(() => {
                        changeTable(type);
                    }, 100);
                }
            });
        });
    }

    const options = `<div class="mb-6 flex items-center" style="justify-content: flex-start;">
				        <div>Enter自动查询：
				        	<select id="enterChoice" class="border rounded-md px-4 py-2 mr-4 flex-1" name="enterChoice">
				        		<option value="0">模拟洗词条</option>
				        		<option value="1">模拟打装备</option>
				        	</select>
				        </div>
				        <div>耳坠：
				        	<select id="earbobChoice" class="border rounded-md px-4 py-2 mr-4 flex-1" name="earbobChoice">
					        	<option value="0">DP+1200</option>
					        	<option value="1">DP+1200 智慧40+</option>
					        	<option value="2">DP+1200 智慧48</option>
				       			<option value="3">全部显示</option>
				        	</select>
				        </div>
				        <div>手环：
					        <select id="wristbandChoice" class="border rounded-md px-4 py-2 mr-4 flex-1" name="wristbandChoice">
						        <option value="0">80+</option>
						        <option value="1">双45</option>
						        <option value="2">全部显示</option>
						        <option value="3">+3手环(导出显示)</option>
					        </select>
				        </div>
				        <div class="mr-4">项链：2 行；耳坠：3 行；手环：4 行</div>
				        <button id="exportButton" class="bg-primary text-primary-foreground rounded-md px-4 py-2 mr-4 hover:bg-primary/90">导出200条</button>
		        	</div>`;

    // 新增操作栏
    setTimeout(() => {
        document.querySelector("body > div").insertBefore(document.createRange().createContextualFragment(options).firstChild, document.querySelector("body > div > div.overflow-x-auto"));

        const enterSelect = document.getElementById("enterChoice");
        const earbobSelect = document.getElementById("earbobChoice");
        const wristbandSelect = document.getElementById("wristbandChoice");
        const exportButton = document.getElementById("exportButton");

        enterSelect.selectedIndex = GM_getValue("enterChoice") ?? "1";
        earbobSelect.selectedIndex = GM_getValue("earbobChoice") ?? "1";
        wristbandSelect.selectedIndex = GM_getValue("wristbandChoice") ?? "0";

        enterSelect.addEventListener("change", (e) => {
            changeSelect(e);
        });
        earbobSelect.addEventListener("change", (e) => {
            changeSelect(e);
        });
        wristbandSelect.addEventListener("change", (e) => {
            changeSelect(e);
        });
        exportButton.addEventListener("click", (e) => {
            toExcel(GM_getValue("currentType"), 200);
        })
    }, 100);

    const _模拟洗词条 = document.querySelector(
        "body > div > div.mb-6.flex.items-center > button:nth-child(3)"
    );
    const _模拟洗装备 = document.querySelector(
        "body > div > div.mb-6.flex.items-center > button:nth-child(4)"
    );
    document.onkeydown = (event) => {
        event = event ?? window.event;
        if (event?.keyCode == 13) {
            if ((GM_getValue("enterChoice") ?? "1") === "1") {
                _模拟洗装备.click();
            } else {
                _模拟洗词条.click();
            }
        }
    };
})();

async function toExcel(type, nums) {
    if (!type) {
        alert("请先查询一遍，再导出，导出将花费15秒");
        return;
    }
    let currentSeed = input1.value;
    let currentIndex = input2.value;
    const sheetDatas = [];
    const turns = nums / 50;
    for (let i = 0; i < turns; i++) {
        const data = await fetch("https://hbrapi.fuyumi.xyz/api/" + type + "?_seed=" + input1.value + "&_index=" + currentIndex).then((response) => response.json());
        for (let j = 0; j < 50; j++) {
            const row = [++currentIndex, ...(data[j].split("/"))];
            sheetDatas.push(row);
        }
        await sleep(1000);
    }
    const len = sheetDatas[0].length;
    sheetDatas.unshift(Array.from({length: len}).map((item, index) => {
        return index
    }));

    // 执行导出
    // const XLSX = await import("https://cdn.sheetjs.com/xlsx-0.20.2/package/xlsx.mjs");
    const XLSX = require("xlsx-js-style");
    const wb= XLSX.utils.book_new()

    const sheetName = `${input1.value}-${currentIndex}`;
    const sheet = XLSX.utils.aoa_to_sheet(sheetDatas);
    changeSheet(type, sheetDatas, sheet);
    sheet['!cols'] = len === 2 ? [{ wch: 10 }, { wch: 20 }] : [{ wch: 5 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 30 }, { wch: 15 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, sheet, sheetName)
    XLSX.writeFile(wb, (type === "RandomMainAbility" ? "饰品" : "属性") + `-${currentSeed}.xlsx`);
}

function changeSelect(event) {
    const id = event.srcElement.id;
    const value = event.srcElement.value;
    GM_setValue(id, value);
    if (id === "enterChoice") {
        return;
    }
    clearAllStyle(oldTbody, "RandomMainAbility");
    changeTable("RandomMainAbility");
}

const input1 = document.querySelector(
    "body > div > div.mb-6.flex.items-center > input:nth-child(1)"
);
const input2 = document.querySelector(
    "body > div > div.mb-6.flex.items-center > input:nth-child(2)"
);
const tr = document.querySelector(
    "body > div > div.overflow-x-auto > table > thead > tr"
);
const oldTbody = document.querySelector(
    "body > div > div.overflow-x-auto > table > tbody"
);

function changeTable(type) {
    if (type === "ChangeAbility") {
        return;
    }
    const tbody = document.querySelector("body > div > div.overflow-x-auto > table > tbody");
    if (tbody.childElementCount < 2) {
        return;
    }
    changeEarbobTable(tbody);
    changeWristbandTable(tbody);
}

function changeEarbobTable(tbody) {
    // 获取tbody中所有的子元素
    // 每三个为一组，一共有50行，7列
    // 对于第一行第二列，第二行第三列，第三行第四列，都将这一个子元素都设置为红色，不满足3的倍数列不做处理
    if (tbody) {
        const children = tbody.children;
        const earbob = 3; // 耳坠占用词条数
        const columnIndexToColor = [1, 2, 3]; // 第一行第二列，第二行第三列，第三行第四列
        const earbobChoice = GM_getValue("earbobChoice") ?? "1";

        for (let i = 0; i < children.length; i++) {
            const tr = children[i];
            const _DP = children[i - 1]?.children[2].innerText.split("+")[1]?.trim();
            const _智慧 = children[i].children[3].innerText.split("+")[1]?.trim();
            if ((earbobChoice === "0" || earbobChoice === "3") && i > 0 && _DP === "1200") {
                // tr.children[3].style.backgroundColor = "yellow";
                children[i - 1].children[2].style.backgroundColor = "yellow";
            }
            if ((earbobChoice === "1" || earbobChoice === "3") && i > 0 && _智慧.startsWith("4") && _DP === "1200") {
                tr.children[3].style.backgroundColor = "yellow";
                children[i - 1].children[2].style.backgroundColor = "yellow";
            }
            if ((earbobChoice !== "0") && i > 0 && _智慧 === "48" && _DP === "1200") {
                tr.children[3].style.backgroundColor = "red";
                children[i - 1].children[2].style.backgroundColor = "red";
            }
            const columnIndex = (i % earbob); // 计算当前子元素所在的列索引（从1开始）
            if (!columnIndexToColor[columnIndex]) {
                continue;
            }
            tr.children[columnIndexToColor[columnIndex]].style.color = "red"; // 将指定的列设置为红色
        }
    } else {
        console.log("未找到 tbody 元素");
    }
}

function changeWristbandTable(tbody) {
    // 获取tbody中所有的子元素
    // 每四个为一组，一共有50行，7列
    // 对于第一行第5列，第三行第6列，第四行第7列，都将这一个子元素都设置为蓝色，不满足4的倍数列不做处理
    if (tbody) {
        const children = tbody.children; // 将HTMLCollection转换为数组
        const wristband = 4; // 手环占用词条数
        const columnIndexToColor = [4, undefined, 5, 6]; // 第一行第5列，第三行第6列，第四行第7列
        const wristbandChoice = GM_getValue("wristbandChoice") ?? "0";

        for (let i = 0; i < children.length; i++) {
            const tr = children[i];
            const _体力 = children[i - 1]?.children[5].innerText.split("+")[1]?.trim();
            const _精神 = children[i].children[6].innerText.split("+")[1]?.trim();
            if ((wristbandChoice === "0" || wristbandChoice === "2") && i > 0 && (Number.parseInt(_体力) + Number.parseInt(_精神) >= 80)) {
                tr.children[6].style.backgroundColor = "yellow";
                children[i - 1].children[5].style.backgroundColor = "yellow";
            }
            if ((wristbandChoice === "1" || wristbandChoice === "2") && i > 0 && _体力 === "45" && _精神 === "45") {
                tr.children[6].style.backgroundColor = "red";
                children[i - 1].children[5].style.backgroundColor = "red";
            }
            const columnIndex = (i % wristband); // 计算当前子元素所在的列索引（从1开始）
            if (!columnIndexToColor[columnIndex]) {
                continue;
            }
            tr.children[columnIndexToColor[columnIndex]].style.color = "blue"; // 将指定的列设置为红色
        }
    } else {
        console.log("未找到 tbody 元素");
    }
}

function changeSheet(type, data, sheet) {
    if (type === "ChangeAbility") {
        return;
    }

    changeEarbobSheet(data, sheet);
    changeWristbandSheet(data, sheet);
}

const yellowBgStyle = {
    fill: {
        patternType: "solid",
        fgColor: { rgb: "FFFF00" }, // 黄色背景
        bgColor: { rgb: "FFFF00" } // 黄色背景
    }
};

const redBgStyle = {
    fill: {
        patternType: "solid",
        fgColor: { rgb: "FF0000" }, // 红色背景
        bgColor: { rgb: "FF0000" } // 红色背景
    }
};

const blueBgStyle = {
    fill: {
        patternType: "solid",
        fgColor: { rgb: "87CEFA" }, // 天蓝色背景
        bgColor: { rgb: "87CEFA" } // 天蓝色背景
    }
};

const set_cell_style = (sheet, row, col, style) => {
    const letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
    // 返回对应的字母
    const index = `${letters[col]}${row + 1}`;
    sheet[index].s = style;
};

function changeEarbobSheet(data, sheet) {
    // 获取tbody中所有的子元素
    // 每三个为一组，一共有50行，7列
    // 对于第一行第二列，第二行第三列，第三行第四列，都将这一个子元素都设置为红色，不满足3的倍数列不做处理
    if (data) {
        // const earbob = 3; // 耳坠占用词条数
        // const columnIndexToColor = [1, 2, 3]; // 第一行第二列，第二行第三列，第三行第四列
        const earbobChoice = GM_getValue("earbobChoice") ?? "1";

        for (let i = 3; i < data.length; i++) {
            const _DP = data[i - 1][2].split("+")[1]?.trim();
            const _智慧 = data[i][3].split("+")[1]?.trim();
            if ((earbobChoice === "0" || earbobChoice === "3") && _DP === "1200") {
                set_cell_style(sheet, i, 3, yellowBgStyle);
                set_cell_style(sheet, i-1, 2, yellowBgStyle);
            }
            if ((earbobChoice === "1" || earbobChoice === "3") && _智慧.startsWith("4") && _DP === "1200") {
                set_cell_style(sheet, i, 3, yellowBgStyle);
                set_cell_style(sheet, i-1, 2, yellowBgStyle);
            }
            if ((earbobChoice !== "0") && _智慧 === "48" && _DP === "1200") {
                set_cell_style(sheet, i, 3, redBgStyle);
                set_cell_style(sheet, i-1, 2, redBgStyle);
            }
            // const columnIndex = (i % earbob); // 计算当前子元素所在的列索引（从1开始）
            // if (!columnIndexToColor[columnIndex]) {
            //     continue;
            // }
            // tr.children[columnIndexToColor[columnIndex]].style.color = "red"; // 将指定的列设置为红色
        }
    } else {
        console.log("未找到 data");
    }
}

function changeWristbandSheet(data, sheet) {
    // 获取tbody中所有的子元素
    // 每四个为一组，一共有50行，7列
    // 对于第一行第5列，第三行第6列，第四行第7列，都将这一个子元素都设置为蓝色，不满足4的倍数列不做处理
    if (data) {
        // const wristband = 4; // 手环占用词条数
        // const columnIndexToColor = [4, undefined, 5, 6]; // 第一行第5列，第三行第6列，第四行第7列
        const wristbandChoice = GM_getValue("wristbandChoice") ?? "0";

        for (let i = 4; i < data.length; i++) {
            const _体力 = data[i - 1][5].split("+")[1]?.trim();
            const _精神 = data[i][6].split("+")[1]?.trim();
            const _通常攻击攻击力 = data[i - 3][4].split("+")[1]?.trim();
            // sheet.set_cell_style(1, 1, {font_color: "red"})
            // # 设置背景颜色
            // sheet.set_cell_style(1, 1, {background_color: "blue"})
            if ((wristbandChoice === "0" || wristbandChoice === "2") && (Number.parseInt(_体力) + Number.parseInt(_精神) >= 80)) {
                set_cell_style(sheet, i, 6, yellowBgStyle);
                set_cell_style(sheet, i-1, 5, yellowBgStyle);
            }
            if ((wristbandChoice === "1" || wristbandChoice === "2") && _通常攻击攻击力.startsWith("100") && (Number.parseInt(_体力) + Number.parseInt(_精神) > 80 && (Number.parseInt(_体力) + Number.parseInt(_精神) < 90))) {
                set_cell_style(sheet, i, 6, blueBgStyle);
                set_cell_style(sheet, i-1, 5, blueBgStyle);
                set_cell_style(sheet, i-3, 4, blueBgStyle);
            }
            if ((wristbandChoice === "3" || wristbandChoice === "2") && _体力 === "45" && _精神 === "45") {
                set_cell_style(sheet, i, 6, redBgStyle);
                set_cell_style(sheet, i-1, 5, redBgStyle);
            }
            // const columnIndex = (i % wristband); // 计算当前子元素所在的列索引（从1开始）
            // if (!columnIndexToColor[columnIndex]) {
            //     continue;
            // }
        }
    } else {
        console.log("未找到 data");
    }
}

function clearAllStyle(tbody, type) {
    if (tbody) {
        if (tbody.childElementCount < 2) {
            return;
        }
        const children = tbody.children; // 将HTMLCollection转换为数组
        for (let i = 0; i < children.length; i++) {
            const trList = children[i].children;
            for (let j = 1; j < trList.length; j++) {
                trList[j].style.removeProperty('background-color');
                trList[j].style.color = "black";
            }
        }
    } else {
        console.log("未找到 tbody 元素");
    }
}

const sleep = (delay) => new Promise((resolve) => setTimeout(resolve, delay));