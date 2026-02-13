// ====================================================================
// 05_SHEET_MANAGER.GS - CORRIGIDO  
// Gerenciamento de abas (SEM aba Geral + COM aba Pesquisa)
// ====================================================================

const SheetManager = {
  
  COR_HEADER: '#4285f4',
  
  /**
   * Obtém ou cria aba Pesquisa
   */
  obterAbaPesquisa: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Pesquisa');
    
    if (!sheet) {
      sheet = this.criarAbaPesquisa();
    }
    
    return sheet;
  },
  
  /**
   * Cria aba Pesquisa formatada
   */
  criarAbaPesquisa: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Pesquisa');
    
    const headers = [
      'Nome Empresa',
      'Alias',
      'ID Pedido',
      'Número do Pedido',
      'Data Criação',
      'Data Atualização',
      'Status Pedido',
      'Venda de Produtos',
      'Desconto Concedido',
      'Juros de Venda',
      'Valor Frete',
      'Valor Total Pago',
      "SKU's do Pedido",
      'Nome Cliente'
    ];
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers])
      .setBackground(this.COR_HEADER)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    // Ajustar larguras das colunas
    sheet.setColumnWidth(1, 150);  // Nome Empresa
    sheet.setColumnWidth(2, 120);  // Alias
    sheet.setColumnWidth(3, 100);  // ID Pedido
    sheet.setColumnWidth(4, 120);  // Número do Pedido
    sheet.setColumnWidth(5, 130);  // Data Criação
    sheet.setColumnWidth(6, 130);  // Data Atualização
    sheet.setColumnWidth(7, 120);  // Status Pedido
    sheet.setColumnWidth(8, 130);  // Venda de Produtos
    sheet.setColumnWidth(9, 130);  // Desconto Concedido
    sheet.setColumnWidth(10, 120); // Juros de Venda
    sheet.setColumnWidth(11, 100); // Valor Frete
    sheet.setColumnWidth(12, 130); // Valor Total Pago
    sheet.setColumnWidth(13, 200); // SKU's do Pedido
    sheet.setColumnWidth(14, 200); // Nome Cliente
    
    // Congelar primeira linha
    sheet.setFrozenRows(1);
    
    Logger.registrar('SUCESSO', '', '', 'Aba Pesquisa criada', {});
    
    return sheet;
  },
  
  /**
   * Obtém ou cria aba de Logs
   */
  obterAbaLogs: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Logs');
    
    if (!sheet) {
      sheet = this.criarAbaLogs();
    }
    
    return sheet;
  },
  
  /**
   * Cria aba de Logs formatada
   */
  criarAbaLogs: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Logs');
    
    const headers = [
      'Data/Hora',
      'Tipo',
      'Empresa',
      'Alias',
      'Mensagem',
      'Detalhes'
    ];
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers])
      .setBackground(this.COR_HEADER)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 300);
    sheet.setColumnWidth(6, 200);
    
    sheet.setFrozenRows(1);
    
    return sheet;
  },
  
  /**
   * Obtém ou cria aba de empresa específica
   */
  obterOuCriarAbaEmpresa: function(nomeAba) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(nomeAba);
    
    if (!sheet) {
      sheet = this.criarAbaEmpresa(nomeAba);
    }
    
    return sheet;
  },
  
  /**
   * Cria aba de empresa formatada
   */
  criarAbaEmpresa: function(nomeAba) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet(nomeAba);
    
    const headers = [
      'Data',
      'Nome Empresa',
      'Alias',
      'Venda de Produtos',
      'Desconto Concedido',
      'Juros de Venda'
    ];
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers])
      .setBackground(this.COR_HEADER)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 150);
    
    sheet.setFrozenRows(1);
    
    Logger.registrar('SUCESSO', '', '', `Aba '${nomeAba}' criada`, {});
    
    return sheet;
  },
  
  /**
   * Limpa dados da aba Pesquisa
   */
  limparAbaPesquisa: function() {
    const sheet = this.obterAbaPesquisa();
    const ultimaLinha = sheet.getLastRow();
    
    if (ultimaLinha > 1) {
      sheet.deleteRows(2, ultimaLinha - 1);
      Logger.registrar('INFO', '', '', 'Aba Pesquisa limpa', {});
    }
  },
  
  /**
   * Oculta aba por nome
   */
  ocultarAba: function(nomeAba) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(nomeAba);
    
    if (sheet) {
      sheet.hideSheet();
      Logger.registrar('INFO', '', '', `Aba '${nomeAba}' ocultada`, {});
    }
  },
  
  /**
   * Exibe aba por nome
   */
  exibirAba: function(nomeAba) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(nomeAba);
    
    if (sheet) {
      sheet.showSheet();
      Logger.registrar('INFO', '', '', `Aba '${nomeAba}' exibida`, {});
    }
  },
  
  /**
   * Lista todas as abas da planilha
   */
  listarAbas: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    return sheets.map(sheet => ({
      nome: sheet.getName(),
      oculta: sheet.isSheetHidden()
    }));
  },

};