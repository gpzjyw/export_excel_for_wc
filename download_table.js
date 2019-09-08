
$(document).ready(function() {
  const tableFilterNode = $('#pane-learner .table-filter form>div');

  // 添加导出excel按钮
  tableFilterNode.last().append(
    $('<button style="float: right;">导出excel</button>').on({ click: onClick})
  );
});

// 点击事件
function onClick() {
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
  return dataSource;
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
