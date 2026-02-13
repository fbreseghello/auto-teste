// ====================================================================
// 08_EMAIL_MANAGER.GS
// Gerenciamento de configurações de e-mail mensal
// ====================================================================

const EmailManager = {
  
  NOME_ABA: 'Config_Email',
  
  /**
   * Obtém ou cria aba Config_Email
   */
  obterAba: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.NOME_ABA);
    
    if (!sheet) {
      sheet = this.criarAba();
    }
    
    return sheet;
  },
  
  /**
   * Cria aba Config_Email formatada
   */
  criarAba: function() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(this.NOME_ABA);
  
  const headers = ['Destinatários', 'Assunto', 'Corpo E-mail', 'Enviar E-mail', 'Tipo', 'Quando'];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#4285f4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setColumnWidths(1, 6, [250, 300, 500, 120, 150, 120]);
  sheet.setFrozenRows(1);
  
  // Dropdowns
  const ruleEnviar = SpreadsheetApp.newDataValidation()
    .requireValueInList(['SIM', 'NÃO'], true)
    .setAllowInvalid(false)
    .build();
  
  const ruleTipo = SpreadsheetApp.newDataValidation()
  .requireValueInList(['MENSAL', 'DIÁRIO - SUCESSO', 'DIÁRIO - ERRO', 'PERSONALIZADO'], true)
  .setAllowInvalid(false)
  .build();;
  
  sheet.getRange('D2:D100').setDataValidation(ruleEnviar);
  sheet.getRange('E2:E100').setDataValidation(ruleTipo);
  
  Logger.registrar('INFO', '', '', 'Aba Config_Email criada', {});
  return sheet;
},
  
  /**
   * Lê configuração de e-mail
   */
  lerConfiguracao: function() {
  const sheet = this.obterAba();
  const ultimaLinha = sheet.getLastRow();
  
  if (ultimaLinha < 2) return [];
  
  const dados = sheet.getRange(2, 1, ultimaLinha - 1, 6).getValues();
  const configs = [];
  
  dados.forEach(linha => {
    if (linha[0] && linha[3] === 'SIM') {
      configs.push({
        destinatarios: linha[0],
        assunto: linha[1] || 'Yampi - Notificação',
        corpo: linha[2] || '',
        tipo: linha[4] || 'MENSAL',
        quando: linha[5] ? String(linha[5]).trim() : ''
      });
    }
  });
  
  return configs;
},

deveEnviarHoje: function(tipo, quando, temErro) {
  const hoje = new Date();
  const diaHoje = hoje.getDate();
  
  if (tipo === 'DIÁRIO - SUCESSO') return !temErro;
  if (tipo === 'DIÁRIO - ERRO') return temErro;
  
  if (tipo === 'MENSAL') {
    const diaConfig = parseInt(quando);
    return diaHoje === diaConfig;
  }
  
  if (tipo === 'PERSONALIZADO') {
    const dias = quando.split(',').map(d => parseInt(d.trim()));
    return dias.includes(diaHoje);
  }
  
  return false;
}
};