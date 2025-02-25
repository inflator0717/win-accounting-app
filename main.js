const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');  // 导入 path 模块




let win;

app.whenReady().then(() => {
    win = new BrowserWindow({
      width: 800,
      height: 600,
      webPreferences: {
        nodeIntegration: true, // 允许渲染进程使用 require
        contextIsolation: false, // 关闭上下文隔离
      }
    });
  
    win.loadFile('index.html');
  
    ipcMain.on('add-record', (event, record) => {
      console.log('Received add-record event in main process:', record);
    });
  });




ipcMain.on('add-record', (event, record, accountType) => {
  // 使用 process.resourcesPath 获取资源路径
  const filePath = path.join(process.resourcesPath, 'data.xlsx');
  //const filePath = path.join(__dirname, 'data.xlsx');
  //const filePath = path.join(userDataPath, 'data.xlsx');

  console.log('File will be saved to:', filePath);

  let wb;

  // 检查文件是否存在
  if (fs.existsSync(filePath)) {
    // 文件存在，读取文件
    wb = xlsx.readFile(filePath);
  } else {
    // 文件不存在，创建新文件
    wb = xlsx.utils.book_new();
    const companySheet = xlsx.utils.aoa_to_sheet([["项目类别", "日期", "类型", "用途", "金额", "银行卡", "标签", "经办人"]]); // 公司账目表头
    const familySheet = xlsx.utils.aoa_to_sheet([["项目类别", "日期", "类型", "用途", "金额", "银行卡", "标签", "经办人"]]); // 家庭账目表头
    xlsx.utils.book_append_sheet(wb, companySheet, 'company_records'); // 添加公司账目工作表到工作簿
    xlsx.utils.book_append_sheet(wb, familySheet, 'family_records'); // 添加家庭账目工作表到工作簿
  }

  // 根据选择的账目类型获取工作表
  let ws = wb.Sheets[accountType === 'company' ? 'company_records' : 'family_records'];

  // 如果工作表为空，确保至少有表头
  if (!ws || Object.keys(ws).length === 0) {
    ws = xlsx.utils.aoa_to_sheet([["项目类别", "日期", "类型", "用途", "金额", "银行卡", "标签", "经办人"]]); // 确保表头
    xlsx.utils.book_append_sheet(wb, ws, accountType === 'company' ? 'company_records' : 'family_records'); // 添加工作表到工作簿
  }

  // 添加新记录
  const newRow = [record.category, record.date, record.type, record.use, record.amount, record.card, record.label, record.editor];
  xlsx.utils.sheet_add_aoa(ws, [newRow], { origin: -1 });

  // 保存文件
  xlsx.writeFile(wb, filePath);

  // // 读取所有记录并返回
  // const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
  // records.shift(); // 删除表头
  // event.reply('update-records', records);
  
  // 读取所有记录并返回
  let records = xlsx.utils.sheet_to_json(ws, { header: 1 });
  records.shift(); // 删除表头

  // 按照录入账目的倒序排列
  records = records.reverse();

  event.reply('update-records', records);
});




  ipcMain.on('get-statistics', (event, accountType) => {
    //const filePath = path.join(__dirname, 'data.xlsx');
    const filePath = path.join(process.resourcesPath, 'data.xlsx');
    //const filePath = path.join(userDataPath, 'data.xlsx');
  
    if (fs.existsSync(filePath)) {
      const wb = xlsx.readFile(filePath);
      const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
      const ws = wb.Sheets[sheetName];
      const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
      records.shift(); // 删除表头
  
      const categoryStats = {};
      const cardStats = {};
  
      records.forEach(record => {
        const [category, , type, , amount, card] = record;
        const amt = parseFloat(amount);
  
        if (!categoryStats[category]) {
          categoryStats[category] = { income: 0, expense: 0, total: 0 };
        }
        if (!cardStats[card]) {
          cardStats[card] = { income: 0, expense: 0, total: 0 };
        }
  
        if (type === '收入') {
          categoryStats[category].income += amt;
          cardStats[card].income += amt;
        } else {
          categoryStats[category].expense += amt;
          cardStats[card].expense += amt;
        }
  
        categoryStats[category].total = categoryStats[category].income - categoryStats[category].expense;
        cardStats[card].total = cardStats[card].income - cardStats[card].expense;
      });
      event.reply('statistics-data', { categoryStats, cardStats });
    } else {
      event.reply('statistics-failure', '文件不存在');
    }
  });

  
// 当所有窗口关闭时退出应用
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});




// 按项目类别查询
ipcMain.on('get-categories', (event, accountType) => {
  // 使用 __dirname 获取当前模块的目录名
  //const filePath = path.join(__dirname, 'data.xlsx');
  //const filePath = path.join(userDataPath, 'data.xlsx');
  const filePath = path.join(process.resourcesPath, 'data.xlsx');
  console.log('File will be read from:', filePath);

  if (fs.existsSync(filePath)) {
    const wb = xlsx.readFile(filePath);
    const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
    const ws = wb.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
    records.shift(); // 删除表头

    const categories = [...new Set(records.map(record => record[0]))];
    event.reply('categories-data', categories);
  } else {
    event.reply('categories-failure', '文件不存在');
  }
});





ipcMain.on('get-usage', (event, filter, accountType) => {
  // 使用 __dirname 获取当前模块的目录名
  //const filePath = path.join(__dirname, 'data.xlsx');
  //const filePath = path.join(userDataPath, 'data.xlsx');
  const filePath = path.join(process.resourcesPath, 'data.xlsx');
  console.log('File will be read from:', filePath);

  if (fs.existsSync(filePath)) {
    const wb = xlsx.readFile(filePath);
    const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
    const ws = wb.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
    records.shift(); // 删除表头

    // 过滤出符合条件的记录
    const filteredRecords = records.filter(record => {
      const recordCategory = record[0];
      const recordDate = new Date(record[1]);
      const startDate = new Date(filter.startDate);
      const endDate = new Date(filter.endDate);

      const isCategoryMatch = filter.category === '' || recordCategory === filter.category;
      return isCategoryMatch && recordDate >= startDate && recordDate <= endDate;
    });

    // 提取用途数据
    const usages = [...new Set(filteredRecords.map(record => record[3]))]; // 假设用途在第四列
    event.reply('usage-data', usages);
  } else {
    event.reply('usage-failure', '文件不存在');
  }
});



//获取标签数据
ipcMain.on('get-label', (event, filter, accountType) => {
  // 使用 __dirname 获取当前模块的目录名
  //const filePath = path.join(__dirname, 'data.xlsx');
  //const filePath = path.join(userDataPath, 'data.xlsx');
  const filePath = path.join(process.resourcesPath, 'data.xlsx');
  console.log('File will be read from:', filePath);

  if (fs.existsSync(filePath)) {
    const wb = xlsx.readFile(filePath);
    const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
    const ws = wb.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
    records.shift(); // 删除表头

    // 过滤出符合条件的记录
    const filteredRecords = records.filter(record => {
      const recordCategory = record[0];
      const recordDate = new Date(record[1]);
      const startDate = new Date(filter.startDate);
      const endDate = new Date(filter.endDate);

      const isCategoryMatch = filter.category === '' || recordCategory === filter.category;
      return isCategoryMatch && recordDate >= startDate && recordDate <= endDate;
    });

    // 提取标签数据
    const labels = [...new Set(filteredRecords.map(record => record[6]))]; 
    event.reply('label-data', labels);
  } else {
    event.reply('label-failure', '文件不存在');
  }
});

    //获取经办人数据
    ipcMain.on('get-editor', (event, filter, accountType) => {
      // 使用 __dirname 获取当前模块的目录名
      //const filePath = path.join(__dirname, 'data.xlsx');
      //const filePath = path.join(userDataPath, 'data.xlsx');
      const filePath = path.join(process.resourcesPath, 'data.xlsx');
      console.log('File will be read from:', filePath);
    
      if (fs.existsSync(filePath)) {
        const wb = xlsx.readFile(filePath);
        const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
        const ws = wb.Sheets[sheetName];
        const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
        records.shift(); // 删除表头
    
        // 过滤出符合条件的记录
        const filteredRecords = records.filter(record => {
          const recordCategory = record[0];
          const recordDate = new Date(record[1]);
          const startDate = new Date(filter.startDate);
          const endDate = new Date(filter.endDate);
    
          const isCategoryMatch = filter.category === '' || recordCategory === filter.category;
          return isCategoryMatch && recordDate >= startDate && recordDate <= endDate;
        });

    // 提取标签数据
    const editors = [...new Set(filteredRecords.map(record => record[7]))]; 
    event.reply('editor-data', editors);
  } else {
    event.reply('editor-failure', '文件不存在');
  }
});

  ipcMain.on('search-records-by-category', (event, category,accountType) => {
    //const filePath = path.join(__dirname, 'data.xlsx');
    //const filePath = path.join(app.getAppPath(), 'data.xlsx');  // 使用 app.getAppPath() 获取应用的根路径
    // 使用 process.resourcesPath 获取资源路径
    const filePath = path.join(process.resourcesPath, 'data.xlsx');
    //const filePath = path.join(userDataPath, 'data.xlsx');
    if (fs.existsSync(filePath)) {
      // const wb = xlsx.readFile(filePath);
      // const ws = wb.Sheets['records'];
      // const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
    const wb = xlsx.readFile(filePath);
    const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
    const ws = wb.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
      records.shift(); // 删除表头

      // 过滤出类别匹配的记录
      const filteredRecords = records.filter(record => record[0] === category);
      // 返回所有相关信息
      event.reply('search-results', filteredRecords);
      console.log('search-results', filteredRecords);
    } else {
      event.reply('search-failure', '文件不存在');
    }
  });

  //按日期范围查询
  ipcMain.on('search-records-by-date', (event, { startDate, endDate }, accountType) => {
    //const filePath = path.join(__dirname, 'data.xlsx');
    //const filePath = path.join(userDataPath, 'data.xlsx');
    // 使用 process.resourcesPath 获取资源路径
    const filePath = path.join(process.resourcesPath, 'data.xlsx');
    if (fs.existsSync(filePath)) {
      // const wb = xlsx.readFile(filePath);
      // const ws = wb.Sheets['records'];
      // const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
    const wb = xlsx.readFile(filePath);
    const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
    const ws = wb.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
      records.shift(); // 删除表头

      // 过滤出日期范围内的记录
      const filteredRecords = records.filter(record => {
        const recordDate = new Date(record[1]);
        return recordDate >= new Date(startDate) && recordDate <= new Date(endDate);
      });

      // 返回所有相关信息
      event.reply('search-results-date', filteredRecords);
      console.log('search-results-date', filteredRecords);
    } else {
      event.reply('search-failure', '文件不存在');
    }
  });



  ipcMain.on('search-records-by-category-and-date', (event, { category, startDate, endDate }, accountType) => {
    //const filePath = path.join(__dirname, 'data.xlsx');
    //const filePath = path.join(app.getAppPath(), 'data.xlsx');  // 使用 app.getAppPath() 获取应用的根路径
    // 使用 process.resourcesPath 获取资源路径
    const filePath = path.join(process.resourcesPath, 'data.xlsx');
    //const filePath = path.join(userDataPath, 'data.xlsx');
    if (fs.existsSync(filePath)) {
        const wb = xlsx.readFile(filePath);
        const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
        const ws = wb.Sheets[sheetName];
        const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
        records.shift(); // 删除表头

        // 过滤出类别匹配且日期范围内的记录
        const filteredRecords = records.filter(record => {
            const recordDate = new Date(record[1]);
            const isCategoryMatch = category === '' || record[0] === category;
            return isCategoryMatch && recordDate >= new Date(startDate) && recordDate <= new Date(endDate);
        });

        // 计算统计结果
        const stats = filteredRecords.reduce((acc, record) => {
            const amount = parseFloat(record[4]);
            if (record[2] === '收入') {
                acc.income += amount;
            } else {
                acc.expense += amount;
            }
            acc.total = acc.income - acc.expense;
            return acc;
        }, { income: 0, expense: 0, total: 0 });

        // 返回所有相关信息
        event.reply('detailed-search-results', { records: filteredRecords, stats });
        console.log('detailed-search-results', filteredRecords);
        console.log('stats', stats);
    } else {
        event.reply('search-failure', '文件不存在');
    }
});



  ipcMain.on('search-records-by-category-and-date-and-usage', (event, { category, startDate, endDate, usage }, accountType) => {
    //const filePath = path.join(__dirname, 'data.xlsx');
    //const filePath = path.join(app.getAppPath(), 'data.xlsx');  // 使用 app.getAppPath() 获取应用的根路径
    // 使用 process.resourcesPath 获取资源路径
    const filePath = path.join(process.resourcesPath, 'data.xlsx');
    //const filePath = path.join(userDataPath, 'data.xlsx');
    if (fs.existsSync(filePath)) {
        const wb = xlsx.readFile(filePath);
        const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
        const ws = wb.Sheets[sheetName];
        const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
        records.shift(); // 删除表头

        // 过滤出类别匹配且日期范围内的记录
        const filteredRecords = records.filter(record => {
            const recordDate = new Date(record[1]);
            const isCategoryMatch = category === '' || record[0] === category;
            return isCategoryMatch && record[3] === usage && recordDate >= new Date(startDate) && recordDate <= new Date(endDate);
        });

        // 计算统计结果
        const stats = filteredRecords.reduce((acc, record) => {
            const amount = parseFloat(record[4]);
            if (record[2] === '收入') {
                acc.income += amount;
            } else {
                acc.expense += amount;
            }
            acc.total = acc.income - acc.expense;
            return acc;
        }, { income: 0, expense: 0, total: 0 });

        // 返回所有相关信息
        event.reply('detailed-search-results01', { records: filteredRecords, stats });
        console.log('detailed-search-results01', filteredRecords);
        console.log('stats', stats);
    } else {
        event.reply('search-failure', '文件不存在');
    }
});


  //按项目类别和标签查询
ipcMain.on('search-records-by-category-and-date-and-label', (event, { category, startDate, endDate, label }, accountType) => {
  //const filePath = path.join(__dirname, 'data.xlsx');
  //const filePath = path.join(app.getAppPath(), 'data.xlsx');  // 使用 app.getAppPath() 获取应用的根路径
  // 使用 process.resourcesPath 获取资源路径
  const filePath = path.join(process.resourcesPath, 'data.xlsx');
  //const filePath = path.join(userDataPath, 'data.xlsx');
  if (fs.existsSync(filePath)) {
      const wb = xlsx.readFile(filePath);
      const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
      const ws = wb.Sheets[sheetName];
      const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
      records.shift(); // 删除表头

      // 过滤出类别匹配且日期范围内的记录
      const filteredRecords = records.filter(record => {
          const recordDate = new Date(record[1]);
          const isCategoryMatch = category === '' || record[0] === category;
          return isCategoryMatch && record[6] === label && recordDate >= new Date(startDate) && recordDate <= new Date(endDate);
      });

      // 计算统计结果
      const stats = filteredRecords.reduce((acc, record) => {
          const amount = parseFloat(record[4]);
          if (record[2] === '收入') {
              acc.income += amount;
          } else {
              acc.expense += amount;
          }
          acc.total = acc.income - acc.expense;
          return acc;
      }, { income: 0, expense: 0, total: 0 });

      // 返回所有相关信息
      event.reply('detailed-search-results02', { records: filteredRecords, stats });
      console.log('detailed-search-results02', filteredRecords);
      console.log('stats', stats);
  } else {
      event.reply('search-failure', '文件不存在');
  }
});

  //按项目类别和经办人查询
  ipcMain.on('search-records-by-category-and-date-and-editor', (event, { category, startDate, endDate, editor }, accountType) => {
    //const filePath = path.join(__dirname, 'data.xlsx');
    //const filePath = path.join(app.getAppPath(), 'data.xlsx');  // 使用 app.getAppPath() 获取应用的根路径
    // 使用 process.resourcesPath 获取资源路径
    const filePath = path.join(process.resourcesPath, 'data.xlsx');
    //const filePath = path.join(userDataPath, 'data.xlsx');
    if (fs.existsSync(filePath)) {
        const wb = xlsx.readFile(filePath);
        const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
        const ws = wb.Sheets[sheetName];
        const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
        records.shift(); // 删除表头
  
        // 过滤出类别匹配且日期范围内的记录
        const filteredRecords = records.filter(record => {
            const recordDate = new Date(record[1]);
            const isCategoryMatch = category === '' || record[0] === category;
            return isCategoryMatch && record[7] === editor && recordDate >= new Date(startDate) && recordDate <= new Date(endDate);
        });
  
        // 计算统计结果
        const stats = filteredRecords.reduce((acc, record) => {
            const amount = parseFloat(record[4]);
            if (record[2] === '收入') {
                acc.income += amount;
            } else {
                acc.expense += amount;
            }
            acc.total = acc.income - acc.expense;
            return acc;
        }, { income: 0, expense: 0, total: 0 });
  
        // 返回所有相关信息
        event.reply('detailed-search-results03', { records: filteredRecords, stats });
        console.log('detailed-search-results03', filteredRecords);
        console.log('stats', stats);
    } else {
        event.reply('search-failure', '文件不存在');
    }
  });


  //银行卡查询
  ipcMain.on('get-cards', (event, accountType) => {
    //const filePath = path.join(__dirname, 'data.xlsx');
    //const filePath = path.join(app.getAppPath(), 'data.xlsx');  // 使用 app.getAppPath() 获取应用的根路径
    // 使用 process.resourcesPath 获取资源路径
    const filePath = path.join(process.resourcesPath, 'data.xlsx');
    //const filePath = path.join(userDataPath, 'data.xlsx');
    if (fs.existsSync(filePath)) {
    const wb = xlsx.readFile(filePath);
    const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
    const ws = wb.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
      records.shift(); // 删除表头

      const cards = [...new Set(records.map(record => record[5]))];
      event.reply('cards-data', cards);
    } else {
      event.reply('cards-failure', '文件不存在');
    }
  });



  ipcMain.on('search-records-by-card', (event, { card, startDate, endDate }, accountType) => {
    //const filePath = path.join(__dirname, 'data.xlsx');
    //const filePath = path.join(app.getAppPath(), 'data.xlsx');  // 使用 app.getAppPath() 获取应用的根路径
    // 使用 process.resourcesPath 获取资源路径
    const filePath = path.join(process.resourcesPath, 'data.xlsx');
    //const filePath = path.join(userDataPath, 'data.xlsx');
    if (fs.existsSync(filePath)) {
    const wb = xlsx.readFile(filePath);
    const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
    const ws = wb.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
      records.shift(); // 删除表头

      // 过滤出银行卡匹配且日期范围内的记录
      const filteredRecords = records.filter(record => {
        const recordDate = new Date(record[1]);
        return record[5] === card && recordDate >= new Date(startDate) && recordDate <= new Date(endDate);
      });

      // 计算统计结果
      const stats = filteredRecords.reduce((acc, record) => {
        const amount = parseFloat(record[4]);
        if (record[2] === '收入') {
          acc.income += amount;
        } else {
          acc.expense += amount;
        }
        acc.total = acc.income - acc.expense;
        return acc;
      }, { income: 0, expense: 0, total: 0 });

      // 返回所有相关信息
      event.reply('card-search-results', { records: filteredRecords, stats });
      console.log('card-search-results', filteredRecords);
      console.log('stats', stats);
    } else {
      event.reply('search-failure', '文件不存在');
    }
  });



  // 导出Excel文件
ipcMain.on('show-save-dialog', (event) => {
  dialog.showSaveDialog({
    title: '导出Excel文件',
    defaultPath: '账目信息.xlsx',
    filters: [
      { name: 'Excel文件', extensions: ['xlsx'] }
    ]
  }).then(result => {
    if (!result.canceled) {
      const filePath = result.filePath;
      //const dataFilePath = path.join(__dirname, 'data.xlsx');
      //const dataFilePath = path.join(userDataPath, 'data.xlsx');
      // const dataFilePath = path.join(app.getAppPath(), 'data.xlsx');  // 确保获取正确的路径
      // 使用 process.resourcesPath 获取资源路径
      const dataFilePath = path.join(process.resourcesPath, 'data.xlsx');
      if (fs.existsSync(dataFilePath)) {
        const wb = xlsx.readFile(dataFilePath);

        // 创建一个新的工作簿
        const newWb = xlsx.utils.book_new();

        // 遍历所有工作表并添加到新的工作簿中
        wb.SheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          xlsx.utils.book_append_sheet(newWb, ws, sheetName);
        });

        // 写入文件
        xlsx.writeFile(newWb, filePath);
        event.reply('export-excel-success', filePath);
      } else {
        event.reply('export-excel-failure', '数据文件不存在');
      }
    }
  }).catch(err => {
    console.error('导出Excel文件时出错:', err);
    event.reply('export-excel-failure', '导出Excel文件时出错');
  });
});

  //删除账目
  ipcMain.on('get-records', (event, accountType) => {
    //const filePath = path.join(__dirname, 'data.xlsx');
    //const filePath = path.join(userDataPath, 'data.xlsx');
    //const filePath = path.join(app.getAppPath(), 'data.xlsx'); // 使用 app.getAppPath() 获取应用的根路径
    // 使用 process.resourcesPath 获取资源路径
    const filePath = path.join(process.resourcesPath, 'data.xlsx');
    if (fs.existsSync(filePath)) {
    const wb = xlsx.readFile(filePath);
    const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
    const ws = wb.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
      records.shift(); // 删除表头

      event.reply('records-data', records);
    } else {
      event.reply('records-failure', '数据文件不存在');
    }
  });





  // 删除账目
ipcMain.on('delete-record', (event, recordIndex, accountType) => {
  //const filePath = path.join(__dirname, 'data.xlsx'); // 使用合适的路径获取 data.xlsx 文件
  const filePath = path.join(process.resourcesPath, 'data.xlsx');
  if (fs.existsSync(filePath)) {
    const wb = xlsx.readFile(filePath);
    const sheetName = accountType === 'company' ? 'company_records' : 'family_records';
    const ws = wb.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(ws, { header: 1 });
    records.shift(); // 删除表头

    // 将倒序索引转换为原始索引
    const originalIndex = records.length - 1 - recordIndex;  // 反转索引，恢复原始顺序

    if (originalIndex >= 0 && originalIndex < records.length) {
      records.splice(originalIndex, 1); // 删除指定索引的记录

      // 重新生成工作表并写回文件
      const newWs = xlsx.utils.aoa_to_sheet([["项目类别", "日期", "类型", "用途", "金额", "银行卡", "标签", "经办人"], ...records]);
      wb.Sheets[sheetName] = newWs; // 更新对应的工作表

      // 写回更新后的 Excel 文件
      xlsx.writeFile(wb, filePath);

      event.reply('delete-record-success');
    } else {
      event.reply('delete-record-failure', '无效的记录索引');
    }
  } else {
    event.reply('delete-record-failure', '数据文件不存在');
  }
});
  

  ipcMain.on('exit-app', () => {
    app.quit();
  });


