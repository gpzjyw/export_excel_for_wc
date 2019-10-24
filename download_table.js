
$(document).ready(function() {
  setTimeout(function() {
    const tableFilterNode = $('#pane-learner .table-filter form>div');
  
    // 添加导出excel按钮
    tableFilterNode.last().append(
      $('<button style="margin-left: 30px;">将当前表格数据导出为excel</button>').on({ click: onClick})
    );
    console.log('the export button has been added');
  }, 500);
});

// 点击事件
function onClick(e) {
  if (e) {
    e.stopPropagation();
    e.preventDefault();
  }
  const dataSource = getData();
  exportExcel(dataSource);
}

// 获取表格数据
function getData() {
  const tableRows = $('#pane-learner .el-table__body .el-table__row');
  const dataSource = [];
  for (let index = 0; index < tableRows.length; index++) {
    const rowVal = $('.cell', tableRows[index]);
    dataSource.push({
      name: rowVal[1].innerText,
      selectScore: rowVal[4].innerText,
      totalScore: rowVal[5].innerText,
    });
  }
  return arrRemoveRepeat(dataSource, 'name');
}

// 数组去重
function arrRemoveRepeat(arr, key) {
  const obj = {};
  return arr.filter(function(item) {
    if (obj[item[key]]) {
      return false;
    }
    obj[item[key]] = true;
    return true;
  });
}

// 制作并导出excel 
function exportExcel(dataSource) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(
    [{ name: '姓名', selectScore: '当日得分', totalScore: '总得分' }, 
      ...dataSource],
    { header: ['name', 'selectScore', 'totalScore'], skipHeader: true },
  )
  XLSX.utils.book_append_sheet(wb, ws, 'sheet');

  XLSX.writeFile(wb, `高三党支部_${moment().format('YYYYMMDD')}.xlsx`);
}
