const AZ = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

function init({ exc, id, container, props, ctx }) {
    //exc('load(["https://z.zccdn.cn/vendor/jspreadsheet/jsuites_v5.css", "https://z.zccdn.cn/vendor/jspreadsheet/jspreadsheet_v10.css", "https://z.zccdn.cn/vendor/jspreadsheet/jspreadsheet_v10.js", "https://z.zccdn.cn/vendor/jspreadsheet/jsuites_v5.js"])', null, () => {
    exc('load(["https://z.zccdn.cn/vendor/jspreadsheet/v4/jsuites.css", "//z.zccdn.cn/vendor/jspreadsheet/v4/jexcel.css", "//z.zccdn.cn/vendor/jspreadsheet/v4/jexcel.js", "//z.zccdn.cn/vendor/jspreadsheet/v4/jsuites.js"])', null, () => {
        let model, present
        let past = []
        let future = []
        let fields = []
        let data = []
        let columns = []
        let meta = {}
        let D = props.data
        let C = props.dcolumns || props.columns || {}
        if (!D) return log("没有数据")
        if (Array.isArray(D)) {
            if (Array.isArray(C)) {
                data = D
                columns = C
            } else {
                fields = typeof C === "object" && Object.keys(C).length ? Object.keys(C) : Object.keys(D[0])
                D.forEach((o, j) => {
                    data.push(new Array(fields.length))
                })
                fields.forEach((k, j) => {
                    columns.push(Object.assign({ title: k, width: Math.max(getLen(D[0][k]), 5) * 13 }, C[k]))
                    D.forEach((o, i) => {
                        data[i][j] = o[k]
                    })
                })
            }
        } else if (D.count && D.model) {
            model = D.model
            D = [...D.arr]
            D.forEach((o, i) => {
                meta["A" + (i + 1)] = { _id: o._id }
            })
            fields = Object.keys(C)
            if (fields.length) {
                D.forEach((o, j) => {
                    data.push(new Array(fields.length))
                })
                fields.forEach((k, j) => {
                    columns.push(Object.assign({ title: k }, C[k]))
                    D.forEach((o, i) => {
                        data[i][j] = o.x[k]
                    })
                })
            } else {
                C = {}
                D.forEach((o, j) => {
                    Object.assign(C, o.x)
                })
                fields = Object.keys(C)
                D.forEach((o, j) => {
                    data.push(new Array(fields.length))
                })
                fields.forEach((k, j) => {
                    columns.push({ title: k, width: Math.max(getLen(C[k]), 5) * 13 })
                    D.forEach((o, i) => {
                        data[i][j] = o.x[k]
                    })
                })
            }
        }
        const excel = container.excel = jspreadsheet(container, {
            data,
            columns,
            meta,
            about: false,
            allowDeleteColumn: false,
            allowDeleteRow: false,
            allowInsertColumn: false,
            allowInsertRow: false,
            allowManualInsertColumn: false,
            allowManualInsertRow: false,
            allowRenameColumn: false,
            includeHeadersOnDownload: true,
            defaultColAlign: "left",
            editable: !!model,
            filters: props.filters,
            contextMenu: () => { return false },
            onafterchanges: (instance, arr) => {
                if (!model) return
                if (present) past.push(present)
                present = arr.filter(a => a.newValue !== undefined)
                present.forEach(o => o._id = excel.getMeta("A" + (parseInt(o.y) + 1))._id)
                future = []
                if (props.change) exc(props.change, { ...ctx, $x: present }, () => exc("render()"))
            },
            onundo: (a, b, c) => {
                if (present) future = [present, ...future]
                present = past.pop()
            },
            onredo: () => {
                if (present) past = [...past, present]
                present = future.shift()
            },
        })
        container.getData = excel.getData
        container.export = excel.download
        container.fullscreen = excel.fullscreen
        container.showChanges = () => {
            container.hideChanges
            let changes = [...past]
            if (present) changes.push(present)
            let cell, O = {}
            let id_y = {}
            let meta = excel.getMeta()
            Object.keys(meta).forEach(k => {
                id_y[meta[k]._id] = k.substr(1)
            })
            let y
            changes.forEach(arr => {
                arr.forEach(o => {
                    if (!O[o.x]) O[o.x] = {}
                    y = parseInt(id_y[o._id]) - 1
                    O[o.x][y] ? O[o.x][y].push(o.newValue) : O[o.x][y] = [o.oldValue, o.newValue]
                })
            })
            Object.keys(O).forEach(x => {
                Object.keys(O[x]).forEach(y => {
                    cell = $('#' + id + ' td[data-x="' + x + '"][data-y="' + y + '"]')
                    cell.classList.add("zp136changed")
                    cell.title = "修改历史： " + O[x][y].join(" -> ")
                })
            })
        }
        container.hideChanges = () => {
            $$("#" + id + " .zp136changed").forEach(x => {
                x.classList.remove("zp136changed")
                x.title = ""
            })
        }
        container.saveChangesToDB = exp => {
            if (!model) return
            let changes = [...past]
            if (present) changes.push(present)
            let field
            let O = {}
            changes.forEach(arr => {
                arr.forEach(o => {
                    field = fields[o.x]
                    if (!O[o._id]) O[o._id] = {}
                    O[o._id][field] = o.newValue && C[field].type === "number" ? parseFloat(o.newValue) : o.newValue
                })
            })
            Object.keys(O).forEach(_id => {
                if (Object.values(O[_id]).some(a => a === "")) {
                    let o = { $set: {}, $unset: {} }
                    Object.keys(O[_id]).forEach(k => {
                        O[_id][k] === "" ? o.$unset[k] = "" : o.$set[k] = O[_id][k]
                    })
                    O[_id] = o
                }
            })
            let str = ""
            let R = []
            Object.keys(O).forEach(_id => {
                str += '$' + model + '.modify("' + _id + '", O["' + _id + '"]); R.push($r); '
            })
            exc(str, { O, R }, () => {
                if (typeof exp === "string") exc(exp, { ...ctx, $R: R })
            })
            excel.setData() // null will reload what is in memory
            past = []
            present = undefined
            future = []
        }
    })
}

function getLen(str) { // 计算字符串长度(英文占1个字符，中文汉字占2个字符)
    if (str == null) return 0;
    if (typeof str != "string") str += ""
    return str.replace(/[^\x00-\xff]/g, "xx").length // 把双字节的替换成两个单字节的然后再获得长度
}

const css = `
.jexcel > tbody > tr > td.readonly {
    color: unset !important;
}
.jexcel > thead > tr > td.arrow-down,
.jexcel > tbody > tr > td.jexcel_dropdown {
    background-position: top 50% right -2px !important;
}
.zp136changed  {
    position: relative;
}
.zp136changed:after  {
    content: '';
    position: absolute;
    width: 7px;
    height: 7px;
    top: 0;
    right: 0;
    background: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAYAAADED76LAAAACXBIWXMAAAsTAAALEwEAmpwYAAAFuGlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNS42LWMxNDUgNzkuMTYzNDk5LCAyMDE4LzA4LzEzLTE2OjQwOjIyICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ0MgMjAxOSAoV2luZG93cykiIHhtcDpDcmVhdGVEYXRlPSIyMDE5LTAxLTMxVDE4OjU1OjA4WiIgeG1wOk1ldGFkYXRhRGF0ZT0iMjAxOS0wMS0zMVQxODo1NTowOFoiIHhtcDpNb2RpZnlEYXRlPSIyMDE5LTAxLTMxVDE4OjU1OjA4WiIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDphMTlhZDJmOC1kMDI2LTI1NDItODhjOS1iZTRkYjkyMmQ0MmQiIHhtcE1NOkRvY3VtZW50SUQ9ImFkb2JlOmRvY2lkOnBob3Rvc2hvcDpkOGI5NDUyMS00ZjEwLWQ5NDktYjUwNC0wZmU1N2I3Nzk1MDEiIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDplMzdjYmE1ZS1hYTMwLWNkNDUtYTAyNS1lOWYxZjk2MzUzOGUiIGRjOmZvcm1hdD0iaW1hZ2UvcG5nIiBwaG90b3Nob3A6Q29sb3JNb2RlPSIzIj4gPHhtcE1NOkhpc3Rvcnk+IDxyZGY6U2VxPiA8cmRmOmxpIHN0RXZ0OmFjdGlvbj0iY3JlYXRlZCIgc3RFdnQ6aW5zdGFuY2VJRD0ieG1wLmlpZDplMzdjYmE1ZS1hYTMwLWNkNDUtYTAyNS1lOWYxZjk2MzUzOGUiIHN0RXZ0OndoZW49IjIwMTktMDEtMzFUMTg6NTU6MDhaIiBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZG9iZSBQaG90b3Nob3AgQ0MgMjAxOSAoV2luZG93cykiLz4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOmExOWFkMmY4LWQwMjYtMjU0Mi04OGM5LWJlNGRiOTIyZDQyZCIgc3RFdnQ6d2hlbj0iMjAxOS0wMS0zMVQxODo1NTowOFoiIHN0RXZ0OnNvZnR3YXJlQWdlbnQ9IkFkb2JlIFBob3Rvc2hvcCBDQyAyMDE5IChXaW5kb3dzKSIgc3RFdnQ6Y2hhbmdlZD0iLyIvPiA8L3JkZjpTZXE+IDwveG1wTU06SGlzdG9yeT4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz4En6MDAAAAX0lEQVQYlX3KOw6AIBBAwS32RpJADXfx0pTET+ERZJ8F8RODFtONsG0QAoh0CSDM82dqodaBdQXnfoLZQM7gPai+wjNNE8R4pTuAYNZSKZASqL7CMy0LxNgJp30fKYUDi3+vIqb/+rUAAAAASUVORK5CYII=);
    background-repeat: no-repeat;
    background-position: top right;
}`


$plugin({
    id: "zp136",
    props: [{
        prop: "data",
        type: "text",
        label: "数据",
        ph: "用小括号括起来"
    }, {
        prop: "dcolumns",
        type: "text",
        label: "动态列属性",
        ph: "用小括号括起来，优先级高于静态列属性"
    }, {
        prop: "columns",
        type: "json",
        label: "静态列属性"
    }, {
        prop: "filters",
        type: "switch",
        label: "可过滤"
    }, {
        prop: "change",
        type: "exp",
        label: "change"
    }],
    init,
    css
})

/*
https://github.com/jspreadsheet/ce


var data = [
    ['Jazz', 'Honda', '2019-02-12', '', true, '$ 2.000,00', '#777700'],
    ['Civic', 'Honda', '2018-07-11', '', true, '$ 4.000,01', '#007777'],
]
const columns = [
    { type: 'text', title: 'Car', width: 120 },
    { type: 'number', title: 'Price', width: 100, mask: '$ #.##,00', decimal: ',' },
    { type: 'dropdown', title: 'Make', width: 200, source: ["Alfa Romeo", "Audi", "Bmw"] },
    { type: 'checkbox', title: 'Stock', width: 80 },
    { type: 'image', title: 'Photo', width: 120 },
    { type: 'color', width: 100, render: 'square', }
    { type: 'calendar', title: 'Available', width: 200 },
]

{
    align: "center"
    "readOnly": false
    allowEmpty: false
    editor: null
    options: []
    source: []
    title: "类别"
    type: "text"
    width: 100
    mask:'#.##',
}

in numeric, you must define a mask in properties of column for force users to input only numeric
mask:'#,#', decimal:"," // For float
mask:'[-]#' // For int with negative sign

[{
    type: 'dropdown',
    title: 'Working hours',
    source: ['Morning', 'Afternoon', 'Evening']
}, {
    type: 'dropdown',
    title: 'Category',
    width: '300',
    source: [{ 'id': '1', 'name': 'Fruits' }, { 'id': '2', 'name': 'Legumes' }, { 'id': '3', 'name': 'General Food' }]
}, {
    type: 'dropdown',
    title: 'Category',
    source: [
        { id: '1', name: 'Paulo', image: '/templates/default/img/1.jpg', title: 'Admin', group: 'Secretary' },
        { id: '2', name: 'Cosme Sergio', image: '/templates/default/img/2.jpg', title: 'Teacher', group: 'Docent' },
        { id: '3', name: 'Rose Mary', image: '/templates/default/img/3.png', title: 'Teacher', group: 'Docent' },
        { id: '4', name: 'Fernanda', image: '/templates/default/img/3.png', title: 'Admin', group: 'Secretary' },
        { id: '5', name: 'Roger', image: '/templates/default/img/3.png', title: 'Teacher', group: 'Docent' },
    ]
}]

*/