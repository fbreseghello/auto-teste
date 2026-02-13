// ====================================================================
// 01_YAMPI_API.GS - VERSAO COMPLETA COM CORRECAO DE DATAS
// ====================================================================

const StateManager = {
  
  salvarProgresso: function(alias, periodoOriginal, ultimaDataCompleta, estrategia, totalPaginas) {
    try {
      const estado = {
        alias: alias,
        periodoInicio: Utilitarios.normalizarDataParaAPI(periodoOriginal.inicio),
        periodoFim: Utilitarios.normalizarDataParaAPI(periodoOriginal.fim),
        ultimaDataCompleta: Utilitarios.normalizarDataParaAPI(ultimaDataCompleta),
        estrategia: estrategia,
        totalPaginas: totalPaginas,
        timestamp: new Date().toISOString(),
        status: 'EM_PROGRESSO_TIMEOUT'
      };
      
      const props = PropertiesService.getScriptProperties();
      props.setProperty(`ESTADO_${alias}`, JSON.stringify(estado));
      
      Logger.registrar('INFO', '', alias, 'Progresso salvo', estado);
        
    } catch (erro) {
      Logger.registrar('ERRO', '', alias, 'Erro ao salvar progresso', { erro: erro.toString() });
    }
  },
  
  verificarProgressoPendente: function(alias) {
    try {
      const props = PropertiesService.getScriptProperties();
      const estadoSalvo = props.getProperty(`ESTADO_${alias}`);
      
      if (estadoSalvo) {
        const estado = JSON.parse(estadoSalvo);
        
        // Validar se estado é recente (<24h)
        const timestampEstado = new Date(estado.timestamp).getTime();
        const agora = new Date().getTime();
        const diferencaHoras = (agora - timestampEstado) / (1000 * 60 * 60);
        
        if (diferencaHoras > 24) {
          this.limparProgresso(alias);
          Logger.registrar('INFO', '', alias, 'Estado expirado (>24h) - limpo', {});
          return null;
        }
        
        Logger.registrar('INFO', '', alias, 'Estado pendente encontrado', estado);
        return estado;
      }
      
      return null;
      
    } catch (erro) {
      Logger.registrar('ERRO', '', alias, 'Erro ao verificar progresso', { erro: erro.toString() });
      return null;
    }
  },
  
  limparProgresso: function(alias) {
    try {
      const props = PropertiesService.getScriptProperties();
      props.deleteProperty(`ESTADO_${alias}`);
      Logger.registrar('INFO', '', alias, 'Progresso limpo', {});
    } catch (erro) {
      Logger.registrar('ERRO', '', alias, 'Erro ao limpar progresso', { erro: erro.toString() });
    }
  }
};

const YampiAPI = {
  
  BASE_URL: 'https://api.dooki.com.br/v2/',
  LIMITE_POR_PAGINA: 25,
  DELAY_ENTRE_REQUESTS: 1200, // ✅ Aumentado para 1.2s (era 800ms)
  MAX_REGISTROS_SEGUROS: 9500,
  MAX_TENTATIVAS: 2,
  DATA_MINIMA: new Date(2025, 5, 1),        // 01/06/2025  (mês 5 = junho)
  TEMPO_LIMITE_EXECUCAO: 15 * 60 * 1000,
  TEMPO_INICIO_EXECUCAO: null,

/**
 * Calcula estrategia baseado no total de paginas
 */
calcularEstrategia: function(totalPaginas) {
  if (totalPaginas < 100) {
    return { tipo: 'MENSAL', dias: 30 };
  } else if (totalPaginas < 200) {
    return { tipo: 'QUINZENAL', dias: 15 };
  } else if (totalPaginas < 600) {
    return { tipo: 'SEMANAL', dias: 7 };
  } else {
    return { tipo: 'DIARIO', dias: 1 };
  }
},

/**
 * Consulta inicial para obter total de paginas (sem baixar dados)
 */
consultarTotalPaginas: function(alias, token, secretKey, dataInicio, dataFim) {
  const dataInicioNorm = Utilitarios.normalizarDataParaAPI(dataInicio);
  const dataFimNorm = Utilitarios.normalizarDataParaAPI(dataFim);
  
  const params = {
    page: 1,
    limit: 1, // Apenas 1 registro para pegar meta
    'date': `created_at:${dataInicioNorm}|${dataFimNorm}`,
    'include': 'items'
  };
  
  const statusAtivos = obterStatusParaBusca();
  statusAtivos.forEach((statusId, index) => {
    params[`status_id[${index}]`] = statusId;
  });
  
  try {
    const resultado = this.chamarAPI(alias, token, secretKey, params);
    
    if (resultado.meta && resultado.meta.pagination) {
      return resultado.meta.pagination.total_pages || 0;
    }
    
    return 0;
    
  } catch (erro) {
    Logger.registrar('ERRO', '', alias, `Erro ao consultar total: ${erro.message}`, {});
    return 0;
  }
},

  /**
   * Busca pedidos incremental
   * - ATUALIZADO: busca janela recente (até 7 dias ou mudança de mês) e retorna array de pedidos
   * - PENDENTE / EM_PROGRESSO: processa mês a mês via buscarPedidosAdaptativa e retorna { completo, ultimaData }
   */
  buscarPedidosIncremental: function(config) {
    const { alias, token, secretKey, ultimaData, status } = config;
    const hoje = new Date();

    // ==============================
    // CASO 1: STATUS ATUALIZADO
    // ==============================
    if (status === 'ATUALIZADO') {
      // ✅ IMPORTANTE: Se ultimaData é string, adicionar 'T12:00:00' para evitar problema de timezone
      const dataUltimaAtualizacao = ultimaData 
        ? (typeof ultimaData === 'string' 
            ? new Date(ultimaData + 'T12:00:00') 
            : new Date(ultimaData))
        : new Date(hoje);

      const mesAtual = hoje.getMonth();
      const anoAtual = hoje.getFullYear();
      const mesUltimaData = dataUltimaAtualizacao.getMonth();
      const anoUltimaData = dataUltimaAtualizacao.getFullYear();
      const mudouMes = (anoAtual !== anoUltimaData) || (mesAtual !== mesUltimaData);

      let dataInicial, dataFinal;

      if (mudouMes) {
        // Mudou de mês → processar mês atual completo
        dataInicial = new Date(anoAtual, mesAtual, 1);
        dataFinal   = new Date(hoje);
        Logger.registrar(
          'INFO',
          config.nomeEmpresa,
          alias,
          'MUDANCA DE MES detectada - processando mes atual completo',
          {
            ultimaData: Utilitarios.formatarData(dataUltimaAtualizacao),
            de:        Utilitarios.formatarData(dataInicial),
            ate:       Utilitarios.formatarData(dataFinal)
          }
        );
      } else {
        // Mesmo mês → buscar APENAS a partir do dia seguinte à última atualização
        dataInicial = new Date(dataUltimaAtualizacao);
        dataInicial.setDate(dataInicial.getDate() + 1); // ✅ CORRIGIDO: +1 dia (não reprocessar)
        dataFinal = new Date(hoje);
        
        // Se última data é hoje, não há nada novo para buscar
        if (dataInicial > dataFinal) {
          Logger.registrar(
            'INFO',
            config.nomeEmpresa,
            alias,
            'Empresa já está atualizada para hoje',
            { ultimaData: Utilitarios.formatarData(dataUltimaAtualizacao) }
          );
          return []; // Retorna array vazio
        }
        
        Logger.registrar(
          'INFO',
          config.nomeEmpresa,
          alias,
          'Busca incremental ATUALIZADO (do dia seguinte ate hoje)',
          {
            de:  Utilitarios.formatarData(dataInicial),
            ate: Utilitarios.formatarData(dataFinal)
          }
        );
      }

      const dataIniStr = Utilitarios.formatarData(dataInicial);
      const dataFimStr = Utilitarios.formatarData(dataFinal);

      // Usa a função de período padrão com paginação
      const pedidos = this.buscarPedidosPorPeriodo(
        alias,
        token,
        secretKey,
        dataIniStr,
        dataFimStr
      );

      Logger.registrar(
        'SUCESSO',
        config.nomeEmpresa,
        alias,
        'Periodo ATUALIZADO processado',
        {
          de:      dataIniStr,
          ate:     dataFimStr,
          pedidos: pedidos.length
        }
      );

      // Para ATUALIZADO, o contrato segue sendo "retornar array de pedidos"
      return pedidos;
    }

    // ==============================
    // CASO 2: STATUS PENDENTE / EM_PROGRESSO
    // ==============================
    Logger.registrar(
      'INFO',
      config.nomeEmpresa,
      alias,
      `Processamento incremental - Status: ${status}`,
      { ultimaData: ultimaData }
    );

    // Definir data de início a partir da última data + 1 dia
    let dataInicioStr;
    if (ultimaData) {
      try {
        // ✅ IMPORTANTE: Se ultimaData é string, adicionar 'T12:00:00' para evitar problema de timezone
        const dataUltima = typeof ultimaData === 'string' 
          ? new Date(ultimaData + 'T12:00:00') 
          : new Date(ultimaData);
        const proximoDia = new Date(dataUltima);
        proximoDia.setDate(proximoDia.getDate() + 1); // +1 dia para não reprocessar
        dataInicioStr = Utilitarios.formatarData(proximoDia);
      } catch (erro) {
        // fallback para DATA_MINIMA global (junho/2025)
        dataInicioStr = Utilitarios.formatarData(this.DATA_MINIMA);
      }
    } else {
      // Sem última data → começa na DATA_MINIMA (junho/2025)
      dataInicioStr = Utilitarios.formatarData(this.DATA_MINIMA);
    }

    // Calcular data final: último dia do mês da dataInicio OU hoje (o que for menor)
    const dataInicioDate = new Date(dataInicioStr + 'T00:00:00');
    const ultimoDiaMes = new Date(
      dataInicioDate.getFullYear(),
      dataInicioDate.getMonth() + 1,
      0
    );
    const dataFinalDate = ultimoDiaMes < hoje ? ultimoDiaMes : hoje;
    const dataFimStr = Utilitarios.formatarData(dataFinalDate);

    Logger.registrar(
      'INFO',
      '',
      alias,
      `Periodo incremental definido`,
      { de: dataInicioStr, ate: dataFimStr }
    );

    // Chama motor adaptativo (ele mesmo processa blocos e chama DataProcessor)
    try {
      const resultado = this.buscarPedidosAdaptativa(
        alias,
        token,
        secretKey,
        dataInicioStr,
        dataFimStr,
        config
      );

      if (resultado.completo) {
        Logger.registrar(
          'SUCESSO',
          '',
          alias,
          'Periodo incremental COMPLETO',
          { de: dataInicioStr, ate: dataFimStr }
        );
        return { completo: true, ultimaData: dataFimStr };
      } else {
        Logger.registrar(
          'INFO',
          '',
          alias,
          'Processamento parcial - continuacao agendada',
          { ultima_data_completa: resultado.ultimaData }
        );
        return { completo: false, ultimaData: resultado.ultimaData };
      }
    } catch (erro) {
      Logger.registrar(
        'ERRO',
        '',
        alias,
        `Erro na busca incremental: ${erro.message}`,
        {}
      );
      // Em caso de erro aqui, não mudamos status nem data; retornamos
      return { completo: false, ultimaData: ultimaData || null };
    }
  },

  /**
   * Busca pedidos por periodo - CORRIGIDO
   */
  buscarPedidosPorPeriodo: function(alias, token, secretKey, dataInicio, dataFim) {
    // No início de buscarPedidosPorPeriodo() adicionar:
const statusAtivos = obterStatusParaBusca();
Logger.registrar('INFO', '', alias, `Status usados: ${statusAtivos.length} total`, {status: statusAtivos});
    
    // CORRECAO: Normalizar datas no inicio
    const dataInicioNorm = Utilitarios.normalizarDataParaAPI(dataInicio);
    const dataFimNorm = Utilitarios.normalizarDataParaAPI(dataFim);
    
    let pagina = 1;
    let todosPedidos = [];
    let totalPaginas = 1;
    let tentativasVazias = 0;
    let totalAcumulado = 0;
    
    do {
      const params = {
        page: pagina,
        limit: this.LIMITE_POR_PAGINA,
        'date': `created_at:${dataInicioNorm}|${dataFimNorm}`,
        'include': 'items'
      };
      
      const statusAtivos = obterStatusParaBusca();
statusAtivos.forEach((statusId, index) => {
  params[`status_id[${index}]`] = statusId;
});

      try {
        const resultado = this.chamarAPI(alias, token, secretKey, params);
        
        if (!resultado.data || resultado.data.length === 0) {
          tentativasVazias++;
          if (tentativasVazias >= 3) break;
        } else {
          tentativasVazias = 0;
          todosPedidos = todosPedidos.concat(resultado.data);
          totalAcumulado += resultado.data.length;
          
          if (totalAcumulado >= this.MAX_REGISTROS_SEGUROS) {
            Logger.registrar('AVISO', '', alias, 
              'Limite de seguranca atingido', { total_acumulado: totalAcumulado });
            break;
          }
        }
        
        if (resultado.meta && resultado.meta.pagination) {
          totalPaginas = resultado.meta.pagination.total_pages || 1;
        }
        
        if (pagina % 10 === 0) {
          Logger.registrar('INFO', '', alias, 
            `Pagina ${pagina}/${totalPaginas}`, 
            { total_acumulado: totalAcumulado, tempo_decorrido: this.getTempoDecorrido() });
        }
        
        pagina++;
        
        // ✅ Pausa adaptativa: mais longa a cada 5 páginas
        if (pagina % 5 === 0) {
          Utilities.sleep(this.DELAY_ENTRE_REQUESTS * 2); // 2.4s
        } else {
          Utilities.sleep(this.DELAY_ENTRE_REQUESTS); // 1.2s
        }
        
      } catch (erro) {
        Logger.registrar('ERRO', '', alias, `Erro na pagina ${pagina}: ${erro.message}`, {});
        throw erro;
      }
      
    } while (pagina <= totalPaginas);
    
    Logger.registrar('SUCESSO', '', alias, 
      `Total de pedidos do periodo: ${todosPedidos.length}`, {});
    
    return todosPedidos;
  },

  /**
   * Busca pedidos com estratégia adaptativa por blocos de dias
   * - Respeita uma janela máxima de ~15 minutos (TEMPO_LIMITE_EXECUCAO)
   * - Salva progresso por alias e período (StateManager)
   * - Retoma exatamente do dia seguinte ao último bloco completo
   * - Processa cada bloco imediatamente (DataProcessor + Aggregator)
   */
  buscarPedidosAdaptativa: function(alias, token, secretKey, dataInicio, dataFim, config) {
    try {
      const dataInicioNorm = Utilitarios.normalizarDataParaAPI(dataInicio);
      const dataFimNorm    = Utilitarios.normalizarDataParaAPI(dataFim);

      // Iniciar timer global desta execução (janela de 15 minutos configurada em TEMPO_LIMITE_EXECUCAO)
      if (!this.TEMPO_INICIO_EXECUCAO) {
        this.TEMPO_INICIO_EXECUCAO = new Date().getTime();
        Logger.registrar('INFO', '', alias, 'Timer de execucao iniciado', {
          limite_ms: this.TEMPO_LIMITE_EXECUCAO
        });
      }

      // Verificar se há progresso pendente para este alias e período
      const estadoPendente = StateManager.verificarProgressoPendente(alias);
      let dataAtualInicio, estrategia, totalPaginas;

      if (
        estadoPendente &&
        estadoPendente.periodoInicio === dataInicioNorm &&
        estadoPendente.periodoFim    === dataFimNorm
      ) {
        // RETOMAR de onde parou: dia seguinte à ultimaDataCompleta
        // ✅ IMPORTANTE: Adicionar 'T12:00:00' para evitar problema de timezone UTC/local
        const ultimaData = new Date(estadoPendente.ultimaDataCompleta + 'T12:00:00');
        ultimaData.setDate(ultimaData.getDate() + 1); // Próximo dia
        dataAtualInicio = Utilitarios.formatarData(ultimaData);

        estrategia = {
          tipo: estadoPendente.estrategia,
          dias: this.getDiasPorEstrategia(estadoPendente.estrategia)
        };
        totalPaginas = estadoPendente.totalPaginas;

        Logger.registrar('INFO', '', alias, 'RETOMANDO processamento', {
          estrategia: estrategia.tipo,
          de:        dataAtualInicio,
          ate:       dataFimNorm
        });
      } else {
        // Novo processamento para este alias/período
        StateManager.limparProgresso(alias);
        dataAtualInicio = dataInicioNorm;

        // Consultar total de páginas uma única vez para o período todo
        Logger.registrar('INFO', '', alias, 'Consultando total de paginas...', {});
        totalPaginas = this.consultarTotalPaginas(
          alias,
          token,
          secretKey,
          dataInicioNorm,
          dataFimNorm
        );
        estrategia = this.calcularEstrategia(totalPaginas);

        Logger.registrar('INFO', '', alias, 'Estrategia definida', {
          total_paginas:  totalPaginas,
          estrategia:     estrategia.tipo,
          dias_por_bloco: estrategia.dias
        });
      }

      // Processar por blocos de dias
      let dataAtual = new Date(dataAtualInicio + 'T00:00:00');
      const dataFinal = new Date(dataFimNorm + 'T23:59:59');

      // Guardamos a última data 100% processada (string yyyy-mm-dd)
      let ultimaDataProcessada = dataAtualInicio;

      while (dataAtual <= dataFinal) {
        // Verificar tempo decorrido antes de iniciar o próximo bloco
        const agora = new Date().getTime();
        const tempoDecorridoMs = agora - this.TEMPO_INICIO_EXECUCAO;

        // Se estamos perto do limite de 15 min, salvar progresso e agendar continuação
        if (tempoDecorridoMs >= (this.TEMPO_LIMITE_EXECUCAO - 60 * 1000)) {
          Logger.registrar('AVISO', '', alias, 'Timeout iminente - salvando progresso', {
            ultima_data_completa: ultimaDataProcessada,
            tempo_decorrido_s:    Math.round(tempoDecorridoMs / 1000)
          });

          StateManager.salvarProgresso(
            alias,
            { inicio: dataInicioNorm, fim: dataFimNorm },
            ultimaDataProcessada,
            estrategia.tipo,
            totalPaginas
          );

          this.agendarContinuacao(alias);
          return { completo: false, ultimaData: ultimaDataProcessada };
        }

        // Calcular fim do bloco
        let dataFimBloco = new Date(dataAtual);
        dataFimBloco.setDate(dataFimBloco.getDate() + estrategia.dias - 1);
        if (dataFimBloco > dataFinal) dataFimBloco = new Date(dataFinal);

        const blocoInicio = Utilitarios.formatarData(dataAtual);
        const blocoFim    = Utilitarios.formatarData(dataFimBloco);

        Logger.registrar('INFO', '', alias,
          `Processando bloco ${estrategia.tipo}: ${blocoInicio} ate ${blocoFim}`, {});

        // Buscar pedidos do bloco
        const pedidosBloco = this.buscarPedidosPorPeriodo(
          alias,
          token,
          secretKey,
          blocoInicio,
          blocoFim
        );

        // Processar bloco imediatamente (upsert seguro nas abas)
        if (pedidosBloco.length > 0) {
          DataProcessor.processarPedidos(pedidosBloco, config);
          Logger.registrar('SUCESSO', '', alias,
            `Bloco processado: ${pedidosBloco.length} pedidos`,
            { de: blocoInicio, ate: blocoFim });
        }

        // Atualizar checkpoint (último dia totalmente processado)
        ultimaDataProcessada = blocoFim;

        // Avançar para próximo bloco
        dataAtual.setDate(dataAtual.getDate() + estrategia.dias);
        
        // ✅ Pausa entre blocos para evitar sobrecarga do Apps Script
        Utilities.sleep(1500); // 1.5s entre blocos
      }

      // Se chegou aqui, todo o período foi processado
      StateManager.limparProgresso(alias);
      Logger.registrar('SUCESSO', '', alias, 'Periodo COMPLETO', {
        estrategia: estrategia.tipo,
        de:         dataInicioNorm,
        ate:        dataFimNorm
      });

      return { completo: true, ultimaData: dataFimNorm };

    } catch (erro) {
      Logger.registrar('ERRO', '', alias, `Erro na estrategia adaptativa: ${erro.message}`, {});
      throw erro;
    }
  },

/**
 * Retorna dias por estrategia
 */
getDiasPorEstrategia: function(tipoEstrategia) {
  const mapa = {
    'MENSAL': 30,
    'QUINZENAL': 15,
    'SEMANAL': 7,
    'DIARIO': 1
  };
  return mapa[tipoEstrategia] || 7;
},

/**
 * Agenda trigger para continuar processamento
 */
agendarContinuacao: function(alias) {
  try {
    // ✅ Limpar estados antigos (> 10 minutos) antes de agendar nova continuação
    const props = PropertiesService.getScriptProperties();
    const allKeys = props.getKeys();
    const agora = new Date().getTime();
    const limiteIdade = 10 * 60 * 1000; // 10 minutos
    
    allKeys
      .filter(key => key.startsWith('ESTADO_'))
      .forEach(key => {
        try {
          const estado = JSON.parse(props.getProperty(key));
          const idade = agora - new Date(estado.timestamp).getTime();
          if (idade > limiteIdade) {
            props.deleteProperty(key);
            Logger.registrar('INFO', '', estado.alias, 'Estado antigo removido antes de agendar continuacao', {
              idade_minutos: Math.round(idade / 60000)
            });
          }
        } catch (e) {
          // Ignorar erros de parse
        }
      });
    
    // Limpar triggers antigos de continuação antes
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      const funcName = trigger.getHandlerFunction();
      if (funcName === 'continuarProcessamentoAdaptativo' || funcName === 'continuarProcessamento') {
        try {
          ScriptApp.deleteTrigger(trigger);
        } catch (e) {
          // Ignorar se já foi deletado
        }
      }
    });
    
    // Criar novo trigger para continuar em 1 minuto
    ScriptApp.newTrigger('continuarProcessamentoAdaptativo')
      .timeBased()
      .after(90 * 1000) // 1.5 minutos
      .create();
    
    Logger.registrar('INFO', '', alias, 'Trigger de continuacao agendado (+1.5min)', {});
  } catch (erro) {
    Logger.registrar('ERRO', '', alias, 'Erro ao agendar continuacao', { erro: erro.toString() });
  }
},

  /**
   * Busca pedidos recentes (ultimos 30 dias) usando updated_at
   */
  buscarPedidosRecentes: function(alias, token, secretKey, dataInicio, dataFim) {
    // CORRECAO: Normalizar datas
    const dataInicioNorm = Utilitarios.normalizarDataParaAPI(dataInicio);
    const dataFimNorm = Utilitarios.normalizarDataParaAPI(dataFim);
    
    let pagina = 1;
    let todosPedidos = [];
    let totalPaginas = 1;
    let tentativasVazias = 0;
    
    do {
      const params = {
        page: pagina,
        limit: this.LIMITE_POR_PAGINA,
        'date': `updated_at:${dataInicioNorm}|${dataFimNorm}`,
        'include': 'items'
      };
      
      const statusAtivos = obterStatusParaBusca();
statusAtivos.forEach((statusId, index) => {
  params[`status_id[${index}]`] = statusId;
});

      try {
        const resultado = this.chamarAPI(alias, token, secretKey, params);
        
        if (!resultado.data || resultado.data.length === 0) {
          tentativasVazias++;
          if (tentativasVazias >= 3) break;
        } else {
          tentativasVazias = 0;
          todosPedidos = todosPedidos.concat(resultado.data);
        }
        
        if (resultado.meta && resultado.meta.pagination) {
          totalPaginas = resultado.meta.pagination.total_pages || 1;
        }
        
        if (pagina % 10 === 0) {
          Logger.registrar('INFO', '', alias, 
            `Pagina ${pagina}/${totalPaginas}`, { total_acumulado: todosPedidos.length });
        }
        
        pagina++;
        Utilities.sleep(this.DELAY_ENTRE_REQUESTS);
        
      } catch (erro) {
        Logger.registrar('ERRO', '', alias, `Erro na pagina ${pagina}: ${erro.message}`, {});
        throw erro;
      }
      
    } while (pagina <= totalPaginas);
    
    Logger.registrar('SUCESSO', '', alias, 
      `Total de atualizacoes: ${todosPedidos.length}`, {});
    
    return todosPedidos;
  },

  /**
   * Chamada HTTP a API Yampi com retry e timeout
   */
  chamarAPI: function(alias, token, secretKey, params) {
    const url = this.construirURL(alias, params);
    
    const options = {
      method: 'get',
      headers: {
        'User-Token': token,
        'User-Secret-Key': secretKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true,
      validateHttpsCertificates: true,
      followRedirects: true,
      // ✅ Timeout de 30 segundos para evitar travamento
      timeout: 30
    };
    
    return this.retryComBackoff(() => {
      try {
        const response = UrlFetchApp.fetch(url, options);
        const statusCode = response.getResponseCode();
        const conteudo = response.getContentText();
        
        if (statusCode !== 200) {
          throw new Error(`Erro ${statusCode}: ${conteudo}`);
        }
        
        return JSON.parse(conteudo);
      } catch (erro) {
        // ✅ Logar erro de timeout especificamente
        if (erro.toString().includes('timeout') || erro.toString().includes('timed out')) {
          Logger.registrar('ERRO', '', alias, 'Timeout na chamada API', { url, erro: erro.toString() });
          throw new Error('API Timeout - tente novamente');
        }
        throw erro;
      }
    });
  },

  /**
   * Constroi URL com query params
   */
  construirURL: function(alias, params) {
    let url = `${this.BASE_URL}${alias}/orders`;
    
    if (params && Object.keys(params).length > 0) {
      const queryString = Object.entries(params)
        .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
        .join('&');
      url += '?' + queryString;
    }
    
    return url;
  },

  /**
   * Retry com backoff exponencial
   */
  retryComBackoff: function(operacao, tentativa = 1) {
    try {
      return operacao();
    } catch (erro) {
      const erroMsg = erro.toString().toLowerCase();
      const isRateLimit = erroMsg.includes('rate limit') || 
                         erroMsg.includes('too many requests') ||
                         erroMsg.includes('429') ||
                         erroMsg.includes('endereco nao disponivel');
      
      if (isRateLimit && tentativa <= this.MAX_TENTATIVAS) {
        const delays = [1000, 3000, 8000];
        const delay = delays[tentativa - 1] || 8000;
        
        Logger.registrar('AVISO', '', '', 
          `Rate limit - tentativa ${tentativa}/${this.MAX_TENTATIVAS}`, 
          { delay_ms: delay, erro: erro.toString() });
        
        Utilities.sleep(delay);
        return this.retryComBackoff(operacao, tentativa + 1);
      } else {
        throw erro;
      }
    }
  },

  tempoDisponivel: function() {
    if (!this.TEMPO_INICIO_EXECUCAO) return true;
    const agora = new Date().getTime();
    const tempoDecorrido = agora - this.TEMPO_INICIO_EXECUCAO;
    const tempoRestante = this.TEMPO_LIMITE_EXECUCAO - tempoDecorrido;
    return tempoRestante > (30 * 1000);
  },

  getTempoDecorrido: function() {
    if (!this.TEMPO_INICIO_EXECUCAO) return '0s';
    const agora = new Date().getTime();
    const decorrido = Math.round((agora - this.TEMPO_INICIO_EXECUCAO) / 1000);
    return `${decorrido}s`;
  },

  testarConexao: function(alias, token, secretKey) {
    const inicio = new Date().getTime();
    try {
      const resultado = this.chamarAPI(alias, token, secretKey, { page: 1, limit: 1 });
      const latencia = new Date().getTime() - inicio;
      return { sucesso: true, latencia: `${latencia}ms`, mensagem: 'API respondendo normalmente' };
    } catch (erro) {
      return { sucesso: false, latencia: '-', mensagem: erro.message };
    }
  },

  descobrirStatusIDs: function(alias, token, secretKey) {
    try {
      const url = `${this.BASE_URL}${alias}/orders/filters`;
      const options = {
        method: 'get',
        headers: {
          'User-Token': token,
          'User-Secret-Key': secretKey,
          'Content-Type': 'application/json'
        }
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const data = JSON.parse(response.getContentText());
      const statusFilter = data.data.find(filter => filter.param === 'status_id');
      
      if (statusFilter) {
        console.log("STATUS ENCONTRADOS:");
        statusFilter.data.forEach(status => {
          console.log(`ID: ${status.id} | Nome: ${status.value}`);
        });
        return statusFilter.data;
      } else {
        console.log("Filtro status_id nao encontrado");
        return null;
      }
    } catch (erro) {
      console.log("Erro:", erro.message);
      return null;
    }
  },

  extrairStatusNome: function(pedido) {
    if (!pedido || !pedido.status) return 'Desconhecido';
    if (pedido.status.data && pedido.status.data.name) {
      return pedido.status.data.name;
    }
    return 'Desconhecido';
  },

  verificarTimeout: function() {
    if (!this.TEMPO_INICIO_EXECUCAO) return false;
    const LIMITE_TEMPO = 5 * 60 * 1000;
    const tempoDecorrido = new Date().getTime() - this.TEMPO_INICIO_EXECUCAO;
    return tempoDecorrido >= LIMITE_TEMPO;
  }

};