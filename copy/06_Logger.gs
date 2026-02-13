// ====================================================================
// 06_LOGGER.GS
// Sistema de logs estruturado
// ====================================================================

const Logger = {
  
  /**
   * Registra log na aba Logs
   */
  registrar: function(tipo, empresa, alias, mensagem, detalhes = {}) {
    try {
      const sheet = SheetManager.obterAbaLogs();
      const timestamp = Utilitarios.formatarTimestamp(new Date());
      const detalhesJson = JSON.stringify(detalhes);
      
      const linha = [
        timestamp,
        tipo,
        empresa || '',
        alias || '',
        mensagem,
        detalhesJson
      ];
      
      const ultimaLinha = sheet.getLastRow();
      const range = sheet.getRange(ultimaLinha + 1, 1, 1, 6);
      range.setValues([linha]);
      
      this.aplicarCorTipo(range.getCell(1, 2), tipo);
      
    } catch (erro) {
      console.error('Erro ao registrar log:', erro);
    }
  },
  
  /**
   * Aplica cor à célula de tipo
   */
  aplicarCorTipo: function(celula, tipo) {
    switch(tipo) {
      case 'ERRO':
        //celula.setBackground('#F4CCCC').setFontColor('#CC0000');
        break;
      case 'AVISO':
        //celula.setBackground('#FFF2CC').setFontColor('#F57C00');
        break;
      case 'SUCESSO':
        //celula.setBackground('#D9EAD3').setFontColor('#38761D');
        break;
      case 'INFO':
        //celula.setBackground('#CFE2F3').setFontColor('#1155CC');
        break;
      case 'INICIO':
        //celula.setBackground('#D0E0E3').setFontColor('#073763');
        break;
    }
  },
  
  /**
   * Obtém último log (para streaming)
   */
  obterUltimoLog: function() {
    try {
      const sheet = SheetManager.obterAbaLogs();
      const ultimaLinha = sheet.getLastRow();
      
      if (ultimaLinha <= 1) return null;
      
      const dados = sheet.getRange(ultimaLinha, 1, 1, 6).getValues()[0];
      
      return {
        timestamp: dados[0],
        tipo: dados[1],
        empresa: dados[2],
        alias: dados[3],
        mensagem: dados[4],
        detalhes: dados[5]
      };
      
    } catch (erro) {
      return null;
    }
  }
};