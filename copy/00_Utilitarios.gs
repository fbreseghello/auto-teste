// ====================================================================
// 00_UTILITARIOS.GS - VERSAO COMPLETA COM CORRECAO
// ====================================================================

const Utilitarios = {
  
  /**
   * Formata data para padrao YYYY-MM-DD
   */
  formatarData: function(data) {
    if (!(data instanceof Date)) {
      data = new Date(data);
    }
    const ano = data.getFullYear();
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const dia = String(data.getDate()).padStart(2, '0');
    return `${ano}-${mes}-${dia}`;
  },
  
  /**
   * NOVA FUNCAO: Normaliza qualquer formato de data para YYYY-MM-DD
   * Aceita: Date, ISO string, YYYY-MM-DD, timestamps
   * Retorna: string no formato YYYY-MM-DD ou null se invalido
   * 
   * Exemplos:
   *   normalizarDataParaAPI("2025-10-31T03:00:00.000Z") => "2025-10-31"
   *   normalizarDataParaAPI("2025-10-31") => "2025-10-31"
   *   normalizarDataParaAPI(new Date()) => "2025-11-28"
   */
  normalizarDataParaAPI: function(data) {
    if (!data) return null;
    
    // Se ja esta no formato correto YYYY-MM-DD, retorna direto
    if (typeof data === 'string') {
      if (/^\d{4}-\d{2}-\d{2}$/.test(data)) {
        return data;
      }
      // Se contem 'T', e ISO string - extrair apenas a parte da data
      if (data.includes('T')) {
        const partesISO = data.split('T')[0];
        if (/^\d{4}-\d{2}-\d{2}$/.test(partesISO)) {
          return partesISO;
        }
      }
    }
    
    // Converte para Date e formata
    try {
      const d = new Date(data);
      if (isNaN(d.getTime())) {
        console.log('Data invalida:', data);
        return null;
      }
      const ano = d.getFullYear();
      const mes = String(d.getMonth() + 1).padStart(2, '0');
      const dia = String(d.getDate()).padStart(2, '0');
      return `${ano}-${mes}-${dia}`;
    } catch (e) {
      console.log('Erro ao normalizar data:', data, e);
      return null;
    }
  },
  
  /**
   * Formata data para formato ISO (usado na API Yampi)
   */
  formatarDataISO: function(data) {
    if (!(data instanceof Date)) {
      data = new Date(data);
    }
    const ano = data.getFullYear();
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const dia = String(data.getDate()).padStart(2, '0');
    return `${ano}-${mes}-${dia}T00:00:00-03:00`;
  },
  
  /**
   * Formata mes/ano para exibicao
   */
  formatarMesAno: function(mes, ano) {
    const meses = [
      'Janeiro', 'Fevereiro', 'Marco', 'Abril', 'Maio', 'Junho',
      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    ];
    const mesIndex = typeof mes === 'number' ? mes - 1 : parseInt(mes) - 1;
    if (mesIndex >= 0 && mesIndex < 12) {
      return `${meses[mesIndex]}/${ano}`;
    }
    return `${mes}/${ano}`;
  },

  /**
   * Obter nome do mes anterior
   */
  obterMesAnterior: function() {
    const hoje = new Date();
    const mesPassado = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
    return this.formatarMesAno(mesPassado.getMonth() + 1, mesPassado.getFullYear());
  },

  /**
   * Formata data/hora para timestamp legivel
   */
  formatarTimestamp: function(data) {
    if (!data) data = new Date();
    const d = new Date(data);
    const dia = String(d.getDate()).padStart(2, '0');
    const mes = String(d.getMonth() + 1).padStart(2, '0');
    const ano = d.getFullYear();
    const hora = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');
    const seg = String(d.getSeconds()).padStart(2, '0');
    return `${dia}/${mes}/${ano} ${hora}:${min}:${seg}`;
  },
  
  /**
   * Extrai data de objeto Yampi
   */
  extrairDataYampi: function(dataObj) {
    if (!dataObj) return null;
    if (typeof dataObj === 'string') return dataObj;
    if (dataObj.date) return dataObj.date.split(' ')[0];
    return null;
  },
  
  /**
   * Normaliza status do pedido
   */
  normalizarStatus: function(statusObj) {
    if (!statusObj) return 'Desconhecido';
    if (typeof statusObj === 'string') return statusObj;
    if (statusObj.data && statusObj.data.name) return statusObj.data.name;
    return 'Desconhecido';
  },
  
  /**
   * Concatena SKUs dos itens
   */
  concatenarSKUs: function(items) {
    if (!items || !Array.isArray(items)) return '';
    return items.map(item => item.item_sku || item.sku_id || '').filter(s => s).join(', ');
  },

  /**
   * Extrai nome do cliente
   */
  extrairNomeCliente: function(customerObj) {
    if (!customerObj) return '';
    if (typeof customerObj === 'string') return customerObj;
    if (customerObj.data && customerObj.data.name) return customerObj.data.name;
    return '';
  },
  
  /**
   * Valida se e mes atual
   */
  ehMesAtual: function(data) {
    const d = new Date(data);
    const hoje = new Date();
    return d.getMonth() === hoje.getMonth() && d.getFullYear() === hoje.getFullYear();
  },
  
  /**
   * Calcula ultimo dia do mes
   */
  ultimoDiaDoMes: function(data) {
    const d = new Date(data);
    return new Date(d.getFullYear(), d.getMonth() + 1, 0);
  },
  
  /**
   * Calcula primeiro dia do proximo mes
   */
  primeiroDiaProximoMes: function(data) {
    const d = new Date(data);
    return new Date(d.getFullYear(), d.getMonth() + 1, 1);
  },

  /**
   * ✅ NOVA FUNÇÃO: Normaliza data para formato aceito pela API (YYYY-MM-DD)
   */
  normalizarDataParaAPI: function(data) {
    if (!data) return null;
    
    try {
      // Se já está no formato correto, retorna
      if (typeof data === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(data)) {
        return data;
      }
      
      // Se é string ISO completa, extrai apenas a data
      if (typeof data === 'string' && data.includes('T')) {
        return data.split('T')[0];
      }
      
      // Se é objeto Date ou string parseável
      const dataObj = new Date(data);
      if (isNaN(dataObj.getTime())) {
        throw new Error('Data inválida');
      }
      
      return this.formatarData(dataObj);
      
    } catch (erro) {
      console.error('Erro ao normalizar data:', erro, 'Data recebida:', data);
      return null;
    }
  },

  /**
   * ✅ NOVA FUNÇÃO: Adiciona dias a uma data e retorna formato YYYY-MM-DD
   */
  adicionarDias: function(data, dias) {
    const dataObj = new Date(data);
    dataObj.setDate(dataObj.getDate() + dias);
    return this.formatarData(dataObj);
  }

};