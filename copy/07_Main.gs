// ====================================================================
// Orquestração principal, menu e triggers
// ====================================================================

/**
 * Cria menu ao abrir planilha
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('➡️ Menu Yampi')
    .addSubMenu(ui.createMenu('📊 Importação')
      .addItem('🏬 Atualizar Todas Empresas', 'atualizacaoManual')
      .addItem('🏢 Atualizar Empresa Específica', 'mostrarSelecaoEmpresa')
      .addItem('🔍 Pesquisa Específica', 'mostrarPesquisaEspecifica'))
    .addSubMenu(ui.createMenu('🔧 Gerenciar Empresas')
      .addItem('➕ Inserir Nova Empresa', 'inserirNovaEmpresa')
      .addItem('🔄 Resetar Todas Empresas', 'resetarTodasEmpresas'))  
    .addSubMenu(ui.createMenu('⚙️ Sistema')
      .addItem('👁️ Mostrar Triggers', 'mostrarTriggers')
      .addItem('🧹 Limpar Logs', 'limparLogs')
      .addItem('🆘 Verificar Saúde da API', 'mostrarSaudeAPI'))
    //.addItem('📧 Configurar E-mail Mensal', 'configurarEmailMensal')
    //.addItem('✉️ Testar Envio de E-mail', 'testarEnvioEmail')
    .addToUi();
}

/**
 * Setup inicial (executar manualmente quando necessário)
 */
function setupInicial() {
  try {
    ConfigManager.obterAba();
    SheetManager.obterAbaLogs();
    SheetManager.obterAbaPesquisa();
    
    configurarTriggersAutomaticos();
    Logger.registrar('INFO', '', '', 'Setup inicial concluído', {});
    
    try {
      SpreadsheetApp.getUi().alert(
        '✅ Estrutura criada!\n\n' +
        'Abas criadas:\n' +
        '• Configuracao\n' +
        '• Pesquisa\n' +
        '• Logs\n\n' +
        'Triggers configurados:\n' +
        '• Sincronização: Diária às 02h\n' +
        '• Limpeza de logs: Diária às 23h\n\n' +
        'Preencha a aba Configuracao com os dados das empresas.'
      );
    } catch (uiError) {
      console.log('✅ Setup concluído! Verifique os logs na planilha.');
    }
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 'Erro no setup inicial', { erro: erro.toString() });
    console.error('Erro no setup:', erro);
    throw erro;
  }
}

/**
 * ✅ MODIFICADO: Configura triggers automaticamente
 */
function configurarTriggersAutomaticos() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      const funcName = trigger.getHandlerFunction();
      if (funcName === 'executarSincronizacao' || 
          funcName === 'iniciarSincronizacaoSemanal' ||
          funcName === 'limparLogsAutomatico' ||
          funcName === 'enviarEmailMensal') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    ScriptApp.newTrigger('iniciarSincronizacaoSemanal')
      .timeBased()
      .everyDays(1)
      .atHour(2)
      .create();
    
    ScriptApp.newTrigger('limparLogsAutomatico')
      .timeBased()
      .everyDays(1)
      .atHour(23)
      .create();
    
    Logger.registrar('SUCESSO', '', '', 
      'Triggers configurados automaticamente', 
      { sincronizacao: 'Diário 02h', limpeza_logs: 'Diário 23h' });
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 
      'Erro ao configurar triggers automáticos', 
      { erro: erro.toString() });
  }
}

/**
 * Mostrar triggers ativos
 */
function mostrarTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const triggersAtivos = [];
    
    triggers.forEach(trigger => {
      const funcName = trigger.getHandlerFunction();
      const source = trigger.getTriggerSource();
      
      if (source === ScriptApp.TriggerSource.CLOCK) {
        const eventType = trigger.getEventType();
        
        let descricao = funcName;
        
        if (eventType === ScriptApp.EventType.WEEK_TIMER) {
          descricao += ' (Semanal)';
        } else if (eventType === ScriptApp.EventType.ON_OPEN) {
          descricao += ' (Ao abrir)';
        } else {
          descricao += ' (Programado)';
        }
        
        triggersAtivos.push(descricao);
      }
    });
    
    let mensagem = '📋 Triggers Ativos:\n\n';
    
    if (triggersAtivos.length === 0) {
      mensagem += 'Nenhum trigger configurado.\n\nExecute "Setup Inicial" para criar triggers automáticos.';
    } else {
      triggersAtivos.forEach(t => {
        mensagem += `• ${t}\n`;
      });
      mensagem += '\n⚙️ Triggers são gerenciados automaticamente.';
    }
    
    SpreadsheetApp.getUi().alert('Triggers do Sistema', mensagem, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (erro) {
    SpreadsheetApp.getUi().alert('Erro ao listar triggers: ' + erro.toString());
  }
}

/**
 * ✅ Inicia sincronização com índice zerado
 */
function iniciarSincronizacaoSemanal() {
  console.log('🔄 Iniciando sincronização semanal...');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('INDICE_SYNC', '0');
  executarSincronizacao();
}

/**
 * ✅ CORRIGIDO: Atualização manual (via menu)
 */
function atualizacaoManual() {
  try {
    console.log('🔄 Iniciando atualização manual...');
    const props = PropertiesService.getScriptProperties();
    props.setProperty('INDICE_SYNC', '0');
        
    executarSincronizacao();
      
    Logger.registrar('INFO', '', '', 'Atualização manual iniciada', {});
    
  } catch (erro) {
    console.error('❌ ERRO na atualização manual:', erro.toString());
    props.deleteProperty('MODO_MANUAL');
    Logger.registrar('ERRO', '', '', 'Erro na atualização manual', 
      { erro: erro.toString() });
  }
}

/**
 * Execução principal da sincronização (LOOP ÚNICO)
 * - Processa múltiplas empresas em loop
 * - Pula empresas já atualizadas hoje
 * - Para PENDENTE/EM_PROGRESSO, respeita resultado.completo da YampiAPI
 * - Cria apenas 1 trigger de continuação se necessário (tempo limite)
 */
function executarSincronizacao() {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(300000)) {
    Logger.registrar('AVISO', '', '', 'Sincronização já em execução', {});
    return;
  }

  try {
    const TEMPO_LIMITE = 15 * 60 * 1000;
    const tempoInicio  = new Date().getTime();
    
    console.log(`⏱️ Tempo limite configurado: ${TEMPO_LIMITE / 1000}s`);

    const props    = PropertiesService.getScriptProperties();
    const empresas = ConfigManager.lerConfiguracao();

    // Filtrar empresas processáveis
    const empresasProcessar = empresas.filter(e =>
      e.status === 'PENDENTE'   ||
      e.status === 'EM_PROGRESSO' ||
      e.status === 'ATUALIZADO'
    );

    if (empresasProcessar.length === 0) {
      Logger.registrar('INFO', '', '', 'Nenhuma empresa para processar', {});
      props.deleteProperty('INDICE_SYNC');
      return;
    }

    // LER ÍNDICE ATUAL
    let indice = parseInt(props.getProperty('INDICE_SYNC') || '0', 10);

    // VERIFICAR SE JÁ TERMINOU O CICLO
    if (indice >= empresasProcessar.length) {
      console.log(`\n========================================`);
      console.log(`✓ SINCRONIZAÇÃO COMPLETA`);
      console.log(`Todas as ${empresasProcessar.length} empresas foram processadas!`);
      console.log(`========================================`);
      
      Logger.registrar('SUCESSO', '', '', 'Todas empresas processadas', {});
      props.deleteProperty('INDICE_SYNC');
      limparTriggersContinuacao();
      return;
    }

    // Log de início (primeira empresa)
    if (indice === 0) {
      console.log(`========================================`);
      console.log(`INÍCIO DA SINCRONIZAÇÃO`);
      console.log(`Total de empresas: ${empresasProcessar.length}`);
      console.log(`========================================`);
      
      Logger.registrar(
        'INICIO',
        '',
        '',
        `Sincronização iniciada (${empresasProcessar.length} empresas)`,
        {}
      );
    }

    // LOOP: Processar empresas enquanto houver tempo
    while (indice < empresasProcessar.length) {
      // Verificar tempo decorrido global da função
      const tempoDecorrido = new Date().getTime() - tempoInicio;
      if (tempoDecorrido > TEMPO_LIMITE) {
        console.log(`\n⏱️ TEMPO LIMITE ATINGIDO (15 minutos)`);
        console.log(`   Empresas processadas: ${indice}/${empresasProcessar.length}`);
        console.log(`   Restantes: ${empresasProcessar.length - indice}`);
        console.log(`   Continuação agendada para 2 minutos...`);
        
        Logger.registrar(
          'AVISO',
          '',
          '',
          'Tempo limite da execucao atingido, agendando continuação',
          {
            empresas_processadas: indice,
            restantes:            empresasProcessar.length - indice
          }
        );

        props.setProperty('INDICE_SYNC', indice.toString());
        // ✅ Limpar triggers antigos antes de criar novo
        limparTriggersContinuacao();
        
        // ✅ Verificar se já existe trigger
        const triggersExistentes = ScriptApp.getProjectTriggers();
        const jaExisteTrigger = triggersExistentes.some(t => 
          t.getHandlerFunction() === 'executarSincronizacao' &&
          t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK
        );
        
        if (!jaExisteTrigger) {
          ScriptApp.newTrigger('executarSincronizacao')
            .timeBased()
            .after(2 * 60 * 1000) // 2 minutos
            .create();
        }

        return;
      }

      const config = empresasProcessar[indice];

      // Pular empresas já atualizadas hoje
      const hojeStr = Utilitarios.formatarData(new Date());
      if (config.status === 'ATUALIZADO' && config.ultimaData === hojeStr) {
        console.log(`   ⊘ ${config.nomeEmpresa}: Já atualizada hoje, pulando...`);
        
        Logger.registrar(
          'INFO',
          config.nomeEmpresa,
          config.alias,
          'Empresa já atualizada hoje, pulando',
          { ultimaData: config.ultimaData }
        );
        indice++;
        continue;
      }

      try {
        // Reset do timer para cada empresa
        YampiAPI.TEMPO_INICIO_EXECUCAO = null;
        
        if (config.status !== 'ATUALIZADO') {
          // Garante que PENDENTE entra como EM_PROGRESSO assim que for processada
          ConfigManager.atualizarStatus(config.alias, 'EM_PROGRESSO', config.ultimaData);
        }

        const tipoProcessamento =
          config.status === 'ATUALIZADO'
            ? 'atualização otimizada (últimos 7 dias)'
            : 'período incremental';

        console.log(`\n[${indice + 1}/${empresasProcessar.length}] Processando: ${config.nomeEmpresa} (${config.alias})`);
        console.log(`   Status: ${config.status} | Última data: ${config.ultimaData}`);
        
        Logger.registrar(
          'INICIO',
          config.nomeEmpresa,
          config.alias,
          `Processando ${tipoProcessamento} (${indice + 1}/${empresasProcessar.length})`,
          { ultimaData: config.ultimaData }
        );

        console.log(`   🔍 Buscando pedidos da API Yampi...`);
        const resultado = YampiAPI.buscarPedidosIncremental(config);
        console.log(`   ✓ API respondeu:`, resultado ? `${Array.isArray(resultado) ? resultado.length : 'objeto'} itens` : 'vazio');

        // ==============================
        // CASO ATUALIZADO
        // ==============================
        if (config.status === 'ATUALIZADO') {
          // Para ATUALIZADO, processar os pedidos retornados
          if (resultado && Array.isArray(resultado) && resultado.length > 0) {
            DataProcessor.processarPedidos(resultado, config);
          }
          
          const hoje = new Date();
          const novaData = Utilitarios.formatarData(hoje);

          ConfigManager.atualizarStatus(config.alias, 'ATUALIZADO', novaData);

          console.log(`   ✓ ${config.nomeEmpresa}: Atualização concluída (${resultado?.length || 0} pedidos)`);
          
          Logger.registrar(
            'SUCESSO',
            config.nomeEmpresa,
            config.alias,
            'Atualização contínua concluída',
            { ultimaData: novaData }
          );
        }

        // ==============================
        // CASO PENDENTE / EM_PROGRESSO
        // ==============================
        else {
          // Se a YampiAPI sinalizar que não concluiu o período, não avança para a próxima empresa
          if (!resultado || resultado.completo !== true) {
            Logger.registrar(
              'INFO',
              config.nomeEmpresa,
              config.alias,
              'Processamento parcial - aguardando continuacao',
              { ultima_data_completa: resultado ? resultado.ultimaData : null }
            );

            // Salva índice atual e encerra (continuacao virá pelo trigger agendado dentro da YampiAPI)
            props.setProperty('INDICE_SYNC', indice.toString());
            return;
          }

          // Período deste mês completo → atualizar status/data
          const dataFinal = new Date(resultado.ultimaData + 'T00:00:00');
          const hoje      = new Date();

          let novoStatus, novaData;
          if (Utilitarios.ehMesAtual(dataFinal)) {
            novoStatus = 'ATUALIZADO';
            novaData   = Utilitarios.formatarData(hoje);
          } else {
            novoStatus = 'EM_PROGRESSO';
            novaData   = Utilitarios.formatarData(dataFinal);
          }

          ConfigManager.atualizarStatus(config.alias, novoStatus, novaData);

          console.log(`   ✓ ${config.nomeEmpresa}: Período processado até ${novaData} - Status: ${novoStatus}`);
          
          Logger.registrar(
            'SUCESSO',
            config.nomeEmpresa,
            config.alias,
            'Período processado',
            { novaData, periodoCompleto: true }
          );
        }
      } catch (erro) {
        console.error(`   ❌ ERRO em ${config.nomeEmpresa}:`, erro.toString());
        
        Logger.registrar(
          'ERRO',
          config.nomeEmpresa,
          config.alias,
          erro.toString(),
          { stack: erro.stack }
        );
        
        // ✅ Verificar se é erro crítico que deve parar tudo
        const erroStr = erro.toString().toLowerCase();
        if (erroStr.includes('service invoked too many times') ||
            erroStr.includes('quota') ||
            erroStr.includes('rate limit')) {
          console.error('⚠️ ERRO CRÍTICO: Limite de API atingido. Parando execução.');
          
          Logger.registrar('ERRO', '', '', 
            'Limite de API atingido - Execução interrompida', 
            { erro: erro.toString() });
          
          // Salvar progresso e parar
          props.setProperty('INDICE_SYNC', indice.toString());
          return;
        }
        
        // Para outros erros, pula a empresa e continua
        console.log(`   ⏭️ Pulando empresa ${config.nomeEmpresa} devido ao erro...`);
      }

      // Próxima empresa
      indice++;
    }

    // TODAS EMPRESAS PROCESSADAS
    Logger.registrar(
      'SUCESSO',
      '',
      '',
      `Todas empresas processadas (${empresasProcessar.length} empresas)`,
      {}
    );

    // Agregação final e e-mail (mantidos como estavam)
    Aggregator.agregarDadosMensais();

    props.deleteProperty('INDICE_SYNC');
    limparTriggersContinuacao();
    enviarEmailMensal();

  } finally {
    lock.releaseLock();
  }
}


/**
 * ✅ Limpar triggers de continuação automaticamente
 */
function limparTriggersContinuacao() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;
    
    triggers.forEach(trigger => {
      const funcName = trigger.getHandlerFunction();
      const source = trigger.getTriggerSource();
      
      if (funcName === 'executarSincronizacao' && source === ScriptApp.TriggerSource.CLOCK) {
        const eventType = trigger.getEventType();
        
        try {
          ScriptApp.deleteTrigger(trigger);
          removidos++;
        } catch (e) {
          // Ignorar erros de triggers já removidos
        }
      }
    });
    
    if (removidos > 0) {
      Logger.registrar('INFO', '', '', 
        `${removidos} trigger(s) de continuação removido(s)`, {});
    }
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 
      'Erro ao limpar triggers de continuação', { erro: erro.toString() });
  }
}

/**
 * Inserir nova empresa (via menu)
 */
function inserirNovaEmpresa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empresas = ConfigManager.lerConfiguracao();
  
  const abasConfiguradas = new Set(empresas.map(e => e.abaDestino));
  const abasCriadas = [];
  
  abasConfiguradas.forEach(nomeAba => {
    if (!ss.getSheetByName(nomeAba)) {
      SheetManager.criarAbaEmpresa(nomeAba);
      abasCriadas.push(nomeAba);
    }
  });
  
  if (abasCriadas.length > 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Abas criadas com sucesso!\n\n' +
      'Novas abas criadas:\n' +
      abasCriadas.map(a => `• ${a}`).join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert(
      'ℹ️ Nenhuma aba nova para criar.\n\n' +
      'Todas as empresas configuradas já possuem suas abas.'
    );
  }
}

/**
 * Limpar logs manualmente
 */
function limparLogs() {
  try {
    const sheet = SheetManager.obterAbaLogs();
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      // ✅ MESMA CORREÇÃO aqui
      const rangeParaLimpar = sheet.getRange(2, 1, lastRow - 1, 6);
      rangeParaLimpar.clearContent();
      
      Logger.registrar('INFO', '', '', 'Logs limpos manualmente', 
        { linhas_limpas: lastRow - 1 });
    } else {
      Logger.registrar('INFO', '', '', 'Nenhum log para limpar', {});
    }
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 'Erro ao limpar logs', 
      { erro: erro.toString() });
  }
}

/**
 * Limpa logs automaticamente (trigger diário às 23h00)
 */
function limparLogsAutomatico() {
  try {
    const sheet = SheetManager.obterAbaLogs();
    const ultimaLinha = sheet.getLastRow();
    
    // Limpar TODAS as linhas de log (exceto cabeçalho)
    if (ultimaLinha > 1) {
      // ✅ CORREÇÃO: Limpar conteúdo ao invés de deletar linhas
      const rangeParaLimpar = sheet.getRange(2, 1, ultimaLinha - 1, 6);
      rangeParaLimpar.clearContent();
      
      Logger.registrar('INFO', '', '', 
        'Limpeza automática de logs executada', 
        { linhas_limpas: ultimaLinha - 1 });
    }
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 
      'Erro na limpeza automática de logs', 
      { erro: erro.toString() });
  }
}

/**
 * Mostrar seleção de empresa específica
 */
function mostrarSelecaoEmpresa() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Interface_Selecao_Empresa')
      .setWidth(550)
      .setHeight(750);
    SpreadsheetApp.getUi().showModalDialog(html, 'Atualizar Empresa Específica');
  } catch (erro) {
    // Fallback se arquivo HTML não existir
    Logger.registrar('ERRO', '', '', 'Erro ao abrir interface de seleção', { erro: erro.toString() });
    
    const empresas = getEmpresasParaSelecao();
    
    if (empresas.empresas.length === 0) {
      SpreadsheetApp.getUi().alert('Nenhuma empresa ativa configurada.');
      return;
    }
    
    const lista = empresas.empresas.map(e => `• ${e.nome_empresa}`).join('\n');
    
    const resposta = SpreadsheetApp.getUi().alert(
      'Empresas Disponíveis',
      `Empresas ativas:\n\n${lista}\n\nUse o menu principal para atualizações automáticas.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Obter lista de empresas para seleção (sem duplicação)
 */
function getEmpresasParaSelecao() {
  const empresas = ConfigManager.lerConfiguracao();
  
  const empresasUnicas = new Map();
  let totalAlias = 0;
  
  empresas.forEach(e => {
    if (e.status !== 'PAUSADO') {
      totalAlias++;
      if (!empresasUnicas.has(e.abaDestino)) {
        empresasUnicas.set(e.abaDestino, e.nomeEmpresa);
      }
    }
  });
  
  const lista = Array.from(empresasUnicas, ([empresa_id, nome_empresa]) => ({ 
    empresa_id, 
    nome_empresa 
  }));
  
  return {
    empresas: lista,
    totalEmpresas: empresasUnicas.size,
    totalAliasAtivos: totalAlias
  };
}

/**
 * Obter empresas para interface HTML (usado pelo modal)
 */
function getAccountsForHtmlInterface() {
  return getEmpresasParaSelecao();
}

/**
 * Obter empresas únicas para seleção (sem duplicação)
 */
function getEmpresasParaPesquisa() {
  const empresas = ConfigManager.lerConfiguracao();
  
  const empresasUnicas = new Map();
  
  empresas.forEach(e => {
    if (e.status !== 'PAUSADO') {
      if (!empresasUnicas.has(e.abaDestino)) {
        empresasUnicas.set(e.abaDestino, {
          empresa_id: e.abaDestino,
          nome_empresa: e.nomeEmpresa
        });
      }
    }
  });
  
  return Array.from(empresasUnicas.values());
}

/**
 * Obter aliases de uma empresa específica
 */
function getAliasesPorEmpresa(empresaId) {
  const empresas = ConfigManager.lerConfiguracao();
  
  const aliases = empresas
    .filter(e => e.abaDestino === empresaId && e.status !== 'PAUSADO')
    .map(e => ({
      alias: e.alias,
      nome_empresa: e.nomeEmpresa
    }));
  
  return aliases;
}

/**
 * Executar pesquisa específica (wrapper para processarPesquisaEspecifica)
 */
function executarPesquisaEspecifica(alias, mesAno) {
  try {
    const resultado = processarPesquisaEspecifica(alias, mesAno);
    return { 
      success: true, 
      registros: resultado.totalPedidos || 0 
    };
  } catch (erro) {
    throw new Error(erro.toString());
  }
}

/**
 * Processar empresa específica via modal
 * @param {string} empresaId - ID da empresa (abaDestino)
 * @param {string} aliasEspecifico - Alias específico (opcional, '' = todos)
 */
function handleAccountSelection(empresaId, aliasEspecifico) {
  const empresas = ConfigManager.lerConfiguracao();
  
  // ✅ Filtrar: se alias específico foi fornecido, processar apenas ele
  let aliasesEmpresa;
  if (aliasEspecifico) {
    aliasesEmpresa = empresas.filter(e => 
      e.abaDestino === empresaId && e.alias === aliasEspecifico
    );
    
    if (aliasesEmpresa.length === 0) {
      throw new Error('Alias específico não encontrado para esta empresa');
    }
    
    Logger.registrar('INICIO', 'Unfair', aliasEspecifico, 
      'Processando via modal - alias específico', {});
  } else {
    // Processar TODOS os aliases da empresa
    aliasesEmpresa = empresas.filter(e => e.abaDestino === empresaId);
    
    if (aliasesEmpresa.length === 0) {
      throw new Error('Empresa não encontrada');
    }
    
    Logger.registrar('INICIO', 'Unfair', empresaId, 
      'Processando via modal - todos os aliases', 
      { total_aliases: aliasesEmpresa.length });
  }
  
  let statsGlobal = { aliasesProcessados: 0, completos: 0, parciais: 0 };

  aliasesEmpresa.forEach(config => {
    try {
      // Reset do timer para cada alias
      YampiAPI.TEMPO_INICIO_EXECUCAO = null;
      
      const status = config.status;
      const ultimaData = config.ultimaData;
      
      statsGlobal.aliasesProcessados++;
      
      Logger.registrar('INICIO', config.nomeEmpresa, config.alias,
        `Processando via modal - status atual: ${status}`,
        { ultimaData });
      
      if (status !== 'ATUALIZADO') {
        ConfigManager.atualizarStatus(config.alias, 'EM_PROGRESSO', ultimaData);
      }
      
      // ✅ USAR A FUNÇÃO QUE REALMENTE EXISTE
      const resultado = YampiAPI.buscarPedidosIncremental(config);
      
      // ==============================
      // CASO ATUALIZADO (retorna array de pedidos)
      // ==============================
      if (status === 'ATUALIZADO') {
        // Para ATUALIZADO, processar os pedidos retornados
        if (resultado && Array.isArray(resultado) && resultado.length > 0) {
          DataProcessor.processarPedidos(resultado, config);
        }
        
        const hoje = new Date();
        const novaData = Utilitarios.formatarData(hoje);
        ConfigManager.atualizarStatus(config.alias, 'ATUALIZADO', novaData);
        
        statsGlobal.completos++;
        
        Logger.registrar('SUCESSO', config.nomeEmpresa, config.alias,
          'Atualização contínua concluída via modal',
          { ultimaData: novaData });
      }
      
      // ==============================
      // CASO PENDENTE/EM_PROGRESSO (retorna {completo, ultimaData})
      // ==============================
      else {
        if (resultado && resultado.completo) {
          statsGlobal.completos++;
          
          const hoje = new Date();
          const dataFinal = new Date(resultado.ultimaData + 'T00:00:00');
          
          let novoStatus, novaData;
          
          if (Utilitarios.ehMesAtual(dataFinal)) {
            novoStatus = 'ATUALIZADO';
            novaData = Utilitarios.formatarData(hoje);
          } else {
            novoStatus = 'EM_PROGRESSO';
            novaData = Utilitarios.formatarData(dataFinal);
          }
          
          ConfigManager.atualizarStatus(config.alias, novoStatus, novaData);
          
          Logger.registrar('SUCESSO', config.nomeEmpresa, config.alias,
            'Período COMPLETO via modal',
            { novoStatus, novaData });
            
        } else {
          statsGlobal.parciais++;
          
          Logger.registrar('INFO', config.nomeEmpresa, config.alias,
            'Processamento PARCIAL - aguardando continuação automática',
            { ultimaDataParcial: resultado && resultado.ultimaData });
        }
      }
      
    } catch (erro) {
      Logger.registrar('ERRO', config.nomeEmpresa, config.alias,
        erro.toString(),
        { stack: erro.stack });
    }
  });
  
  // Agregação no final da execução manual
  Aggregator.agregarDadosMensais();
  
  return {
    success: true,
    stats: statsGlobal,
    empresasProcessadas: aliasesEmpresa.length
  };
}

/**
 * Mostrar interface de pesquisa específica
 */
function mostrarPesquisaEspecifica() {
  const html = HtmlService.createHtmlOutputFromFile('Interface_Pesquisa')
    .setWidth(900)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Pesquisa Específica');
}

/**
 * Processar pesquisa específica
 */
function processarPesquisaEspecifica(alias, mesAno) {
  try {
    const empresas = ConfigManager.lerConfiguracao();
    const config = empresas.find(e => e.alias === alias);
    
    if (!config) {
      throw new Error('Empresa não encontrada');
    }
    
    Logger.registrar('INICIO', config.nomeEmpresa, alias, 
      'Iniciando pesquisa específica', { periodo: mesAno });
    
    const pedidos = DataProcessor.buscarPedidosParaPesquisa(
      alias, config.token, config.secretKey, mesAno
    );
    
    const dadosTransformados = DataProcessor.transformarPedidosParaPesquisa(pedidos, config);
    
    Aggregator.processarDadosPesquisa(dadosTransformados);
    
    Logger.registrar('SUCESSO', config.nomeEmpresa, alias, 
      'Pesquisa específica concluída', { pedidos: pedidos.length });
    
    return {
      success: true,
      totalPedidos: pedidos.length
    };
    
  } catch (erro) {
    Logger.registrar('ERRO', config ? config.nomeEmpresa : '', alias, 
      'Erro na pesquisa específica', { erro: erro.toString() });
    throw erro;
  }
}

/**
 * Limpar aba Pesquisa
 */
function limparAbaPesquisa() {
  SheetManager.limparAbaPesquisa();
}

/**
 * Resetar todas as empresas
 */
function resetarTodasEmpresas() {
  try {
    const empresas = ConfigManager.lerConfiguracao();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Obter abas únicas de destino
    const abasEmpresa = [...new Set(empresas.map(e => e.abaDestino))];
    
    let totalLinhasRemovidas = 0;
    let abasLimpas = 0;
    
    // Limpar cada aba de empresa
    abasEmpresa.forEach(nomeAba => {
      const abaEmpresa = ss.getSheetByName(nomeAba);
      if (abaEmpresa) {
        const lastRow = abaEmpresa.getLastRow();
        if (lastRow > 1) {
          abaEmpresa.deleteRows(2, lastRow - 1);
          totalLinhasRemovidas += (lastRow - 1);
          abasLimpas++;
          Logger.registrar('INFO', nomeAba, '', `${lastRow - 1} linhas deletadas da aba ${nomeAba}`, {});
        }
      }
    });
    
    // Resetar status de todos os aliases
    empresas.forEach(config => {
      ConfigManager.atualizarStatus(config.alias, 'PENDENTE', '2025-07-01');
    });
    
    Logger.registrar('SUCESSO', 'TODAS', '', 
      `Todas empresas resetadas - ${abasLimpas} abas limpas, ${totalLinhasRemovidas} linhas removidas`, 
      { abasLimpas, totalLinhasRemovidas });
    
    // Mostrar resultado
    SpreadsheetApp.getUi().alert(
      'Resetar Todas Empresas',
      `✅ Reset concluído com sucesso!\n\n` +
      `📊 Resumo:\n` +
      `• ${abasLimpas} abas de empresas limpas\n` +
      `• ${totalLinhasRemovidas} linhas de dados removidas\n` +
      `• ${empresas.length} aliases resetados para PENDENTE\n\n` +
      `Todas as empresas estão prontas para nova sincronização.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return { success: true, abasLimpas, totalLinhasRemovidas };
    
  } catch (erro) {
    Logger.registrar('ERRO', 'TODAS', '', 'Erro ao resetar todas empresas', { erro: erro.toString() });
    
    SpreadsheetApp.getUi().alert(
      'Erro no Reset',
      `❌ Erro ao resetar empresas:\n\n${erro.toString()}\n\nVerifique os logs para mais detalhes.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    throw erro;
  }
}

/**
 * Mostrar interface de saúde da API
 */
function mostrarSaudeAPI() {
  const html = HtmlService.createHtmlOutputFromFile('Interface_Saude_API')
    .setWidth(550)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Verificar Saúde da API');
}

/**
 * Verificar saúde da API para todas empresas ativas
 */
function verificarSaudeAPI() {
  const empresas = ConfigManager.lerConfiguracao()
    .filter(e => e.status !== 'PAUSADO');
  
  const resultados = [];
  
  empresas.forEach(config => {
    const resultado = YampiAPI.testarConexao(config.alias, config.token, config.secretKey);
    
    resultados.push({
      empresa: config.nomeEmpresa,
      alias: config.alias,
      status: resultado.sucesso ? '✅ OK' : '❌ ERRO',
      latencia: resultado.latencia,
      mensagem: resultado.mensagem
    });
  });
  
  return resultados;
}

/**
 * Configurar trigger de e-mail mensal
 */
function configurarEmailMensal() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'enviarEmailMensal') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    ScriptApp.newTrigger('enviarEmailMensal')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
    
    Logger.registrar('SUCESSO', '', '', 
      'Trigger de e-mail configurado', 
      { horario: '09:00', frequencia: 'Diário' });
    
    SpreadsheetApp.getUi().alert(
      '✅ E-mail configurado!\n\n' +
      '📧 Verificação diária às 09:00\n' +
      '📝 Configure os envios na aba Config_Email'
    );
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 
      'Erro ao configurar e-mail', 
      { erro: erro.toString() });
  }
}

////////////// FUNÇÕES DE E-MAIL MENSAL /////////////

/**
 * Envia e-mail de lembrete mensal (dia 11) - VERSÃO CORRIGIDA
 */
function enviarEmailMensal() {
  try {
    const configs = EmailManager.lerConfiguracao();
    if (configs.length === 0) return;
    
    const empresas = ConfigManager.lerConfiguracao();
    const temErro = empresas.some(e => e.status === 'ERRO');
    
    configs.forEach(config => {
      if (!EmailManager.deveEnviarHoje(config.tipo, config.quando, temErro)) {
        return;
      }
      
      const mesAnterior = Utilitarios.obterMesAnterior();
      const linkPlanilha = SpreadsheetApp.getActiveSpreadsheet().getUrl();
      
      let corpoFinal = config.corpo
        .replace(/{mesAnterior}/g, mesAnterior)
        .replace(/{linkPlanilha}/g, linkPlanilha);
      
      MailApp.sendEmail({
        to: config.destinatarios,
        subject: config.assunto.replace(/{mesAnterior}/g, mesAnterior),
        htmlBody: corpoFinal
      });
      
      Logger.registrar('SUCESSO', '', '', 
        `E-mail enviado: ${config.tipo}`, 
        { destinatarios: config.destinatarios });
    });
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 
      'Erro ao enviar e-mails', 
      { erro: erro.toString() });
  }
}

/**
 * ✅ Gerar relatório mensal com dados de todas as empresas
 */
function gerarRelatorioMensal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empresas = ConfigManager.lerConfiguracao();
    
    // Obter mês anterior
    const hoje = new Date();     
    const mesAnterior = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);     
    const mesAnteriorStr = Utilitarios.formatarMesAno(mesAnterior.getMonth() + 1, mesAnterior.getFullYear());
    
    const relatorio = {
      mesReferencia: mesAnteriorStr,
      empresas: [],
      totais: {
        receita: 0,
        descontos: 0,
        juros: 0
      }
    };
    
    // Obter empresas únicas (por aba)
    const empresasUnicas = new Map();
    empresas.forEach(e => {
      if (!empresasUnicas.has(e.abaDestino)) {
        empresasUnicas.set(e.abaDestino, {
          nome: e.nomeEmpresa,
          aba: e.abaDestino,
          status: e.status,
          ultimaData: e.ultimaData
        });
      }
    });
    
    // Buscar dados de cada empresa
    empresasUnicas.forEach((empresa, abaDestino) => {
      const abaEmpresa = ss.getSheetByName(abaDestino);
      
      if (abaEmpresa) {
        const dados = buscarDadosMesEmpresa(abaEmpresa, mesAnterior);
        
        if (dados) {
          relatorio.empresas.push({
            nome: empresa.nome,
            receita: dados.receita,
            descontos: dados.descontos,
            juros: dados.juros,
            status: empresa.status
          });
          
          relatorio.totais.receita += dados.receita;
          relatorio.totais.descontos += dados.descontos;
          relatorio.totais.juros += dados.juros;
        }
      }
    });
    
    return relatorio;
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 
      'Erro ao gerar relatório mensal', 
      { erro: erro.toString() });
    return null;
  }
}

/**
 * ✅ Buscar dados de um mês específico de uma empresa
 */
function buscarDadosMesEmpresa(aba, mesReferencia) {
  try {
    const ultimaLinha = aba.getLastRow();
    
    if (ultimaLinha < 2) {
      return null;
    }
    
    const mesReferenciaStr = Utilitarios.formatarData(
      new Date(mesReferencia.getFullYear(), mesReferencia.getMonth(), 1)
    );
    
    // Ler todos os dados
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 6).getValues();
    
    // Buscar linha do mês
    for (let i = 0; i < dados.length; i++) {
      const dataLinha = dados[i][0]; // Coluna A: Data
      
      if (dataLinha && dataLinha.toString() === mesReferenciaStr) {
        return {
          receita: parseFloat(dados[i][3]) || 0,      // Coluna D
          descontos: parseFloat(dados[i][4]) || 0,    // Coluna E
          juros: parseFloat(dados[i][5]) || 0         // Coluna F
        };
      }
    }
    
    return null;
    
  } catch (erro) {
    console.error('Erro ao buscar dados do mês:', erro);
    return null;
  }
}

/**
 * ✅ Montar corpo do e-mail em HTML
 */
function montarCorpoEmail(corpoBase, relatorio) {
  if (!relatorio) {
    return corpoBase || 'Erro ao gerar relatório.';
  }
  
  let html = `
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; }
        .header { background-color: #4285f4; color: white; padding: 20px; text-align: center; }
        .content { padding: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { background-color: #f1f1f1; padding: 10px; text-align: left; border: 1px solid #ddd; }
        td { padding: 10px; border: 1px solid #ddd; }
        .total { font-weight: bold; background-color: #e8f0fe; }
        .footer { margin-top: 20px; color: #666; font-size: 12px; }
      </style>
    </head>
    <body>
      <div class="header">
        <h2>📊 Relatório Mensal Yampi</h2>
        <p>Período: ${relatorio.mesReferencia}</p>
      </div>
      <div class="content">
  `;
  
  // Adicionar corpo personalizado
  if (corpoBase) {
    html += `<p>${corpoBase}</p><hr>`;
  }
  
  // Adicionar tabela de empresas
  html += `
    <h3>Resumo por Empresa</h3>
    <table>
      <thead>
        <tr>
          <th>Empresa</th>
          <th>Status</th>
          <th>Vendas de Produto</th>
          <th>Descontos Concedidos</th>
          <th>Juros de Venda</th>
        </tr>
      </thead>
      <tbody>
  `;
  
  relatorio.empresas.forEach(empresa => {
    html += `
      <tr>
        <td>${empresa.nome}</td>
        <td>${empresa.status}</td>
        <td>R$ ${empresa.receita.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
        <td>R$ ${empresa.descontos.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
        <td>R$ ${empresa.juros.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
      </tr>
    `;
  });
  
  // Linha de totais
  html += `
      <tr class="total">
        <td colspan="2">TOTAL GERAL</td>
        <td>R$ ${relatorio.totais.receita.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
        <td>R$ ${relatorio.totais.descontos.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
        <td>R$ ${relatorio.totais.juros.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
      </tr>
    </tbody>
    </table>
  `;
  
  html += `
        <div class="footer">
          <p>E-mail enviado automaticamente pelo sistema de integração Yampi API.</p>
          <p>Data de envio: ${new Date().toLocaleString('pt-BR')}</p>
        </div>
      </div>
    </body>
    </html>
  `;
  
  return html;
}

/**
 * Continua processamento automatico (legado - redireciona)
 */
function continuarProcessamento() {
  // Redirecionar para função adaptativa
  continuarProcessamentoAdaptativo();
}

/**
 * Limpa triggers de continuacao
 */
function limparTriggersContinuacao() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continuarProcessamento') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Continua processamento adaptativo após timeout
 */
function continuarProcessamentoAdaptativo() {
  try {
    Logger.registrar('INFO', '', '', 'Continuacao automatica iniciada', {});
    
    const props = PropertiesService.getScriptProperties();
    const allKeys = props.getKeys();
    const agora = new Date().getTime();
    const limiteIdade = 10 * 60 * 1000; // 10 minutos
    
    // Carregar estados e filtrar apenas os recentes (últimos 10 minutos)
    const estadosPendentes = allKeys
      .filter(key => key.startsWith('ESTADO_'))
      .map(key => JSON.parse(props.getProperty(key)))
      .filter(estado => {
        const idade = agora - new Date(estado.timestamp).getTime();
        if (idade > limiteIdade) {
          Logger.registrar('INFO', '', estado.alias, 'Estado antigo descartado', { 
            timestamp: estado.timestamp, 
            idade_minutos: Math.round(idade / 60000) 
          });
          props.deleteProperty(`ESTADO_${estado.alias}`);
          return false;
        }
        return true;
      })
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp)); // Mais recente primeiro
    
    if (estadosPendentes.length === 0) {
      Logger.registrar('INFO', '', '', 'Nenhum estado pendente', {});
      limparTriggersContinuacao();
      return;
    }
    
    // Processar estado pendente mais RECENTE
    const estado = estadosPendentes[0];
    const empresas = ConfigManager.lerConfiguracao();
    const config = empresas.find(e => e.alias === estado.alias);
    
    if (!config) {
      StateManager.limparProgresso(estado.alias);
      limparTriggersContinuacao();
      return;
    }
    
    // Reset timer
    YampiAPI.TEMPO_INICIO_EXECUCAO = null;
    
    // Continuar processamento
    const resultado = YampiAPI.buscarPedidosIncremental(config);
    
    // Se completou, atualizar status
    if (resultado.completo) {
      const dataBase = new Date(config.ultimaData);
      const mesProcessado = dataBase.getMonth() + 1;
      const ultimoDiaMes = new Date(dataBase.getFullYear(), mesProcessado + 1, 0);
      const hoje = new Date();
      const dataFinal = ultimoDiaMes < hoje ? ultimoDiaMes : hoje;
      
      let novoStatus, novaData;
      
      if (Utilitarios.ehMesAtual(dataFinal)) {
        novoStatus = 'ATUALIZADO';
        novaData = Utilitarios.formatarData(hoje);
      } else {
        novoStatus = 'EM_PROGRESSO';
        novaData = resultado.ultimaData;
      }
      
      ConfigManager.atualizarStatus(config.alias, novoStatus, novaData);
      
      Logger.registrar('SUCESSO', config.nomeEmpresa, config.alias,
        'Periodo COMPLETO via continuacao', { novaData });
      
      // Limpar TODOS os triggers de continuacao antes de criar novo
      limparTriggersContinuacao();
      
      // Agendar retomada da sincronizacao apenas se veio de execucao automatica (nao modal)
      const indiceSync = props.getProperty('INDICE_SYNC');
      if (indiceSync !== null) {
        const triggersExistentes = ScriptApp.getProjectTriggers();
        const jaExisteTrigger = triggersExistentes.some(t => 
          t.getHandlerFunction() === 'executarSincronizacao' &&
          t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK
        );
        
        if (!jaExisteTrigger) {
          ScriptApp.newTrigger('executarSincronizacao')
            .timeBased()
            .after(10 * 1000)
            .create();
          
          Logger.registrar('INFO', '', '', 'Agendado retomada da sincronizacao', {});
        } else {
          Logger.registrar('INFO', '', '', 'Trigger de sincronizacao ja existe, pulando criacao', {});
        }
      } else {
        Logger.registrar('INFO', '', '', 'Continuacao via modal - sincronizacao geral nao acionada', {});
      }
    }
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 'Erro na continuacao', { erro: erro.toString() });
    limparTriggersContinuacao();
  }
}

/**
 * Processar proximas empresas apos continuacao
 */
function processarProximasEmpresas() {
  try {
    const props = PropertiesService.getScriptProperties();
    const abaDestino = props.getProperty('PROXIMA_EMPRESA_ABA');
    const aliasProcessado = props.getProperty('PROXIMA_EMPRESA_ALIAS');
    
    if (!abaDestino) {
      Logger.registrar('INFO', '', '', 'Nenhuma empresa adicional para processar', {});
      return;
    }
    
    // Limpar flags
    props.deleteProperty('PROXIMA_EMPRESA_ABA');
    props.deleteProperty('PROXIMA_EMPRESA_ALIAS');
    
    const empresas = ConfigManager.lerConfiguracao();
    const aliasesEmpresa = empresas.filter(e => e.abaDestino === abaDestino);
    
    // Encontrar próxima empresa não processada
    let iniciarDe = 0;
    for (let i = 0; i < aliasesEmpresa.length; i++) {
      if (aliasesEmpresa[i].alias === aliasProcessado) {
        iniciarDe = i + 1;
        break;
      }
    }
    
    if (iniciarDe >= aliasesEmpresa.length) {
      Logger.registrar('INFO', '', abaDestino, 'Todas empresas processadas', {});
      
      // Agregação final
      Aggregator.agregarDadosMensais();
      return;
    }
    
    // Processar demais empresas
    Logger.registrar('INFO', '', abaDestino, 
      `Processando empresas restantes (${aliasesEmpresa.length - iniciarDe} restantes)`, {});
    
    for (let i = iniciarDe; i < aliasesEmpresa.length; i++) {
      const config = aliasesEmpresa[i];
      
      try {
        YampiAPI.TEMPO_INICIO_EXECUCAO = null;
        
        if (config.status === 'PENDENTE') {
          ConfigManager.atualizarStatus(config.alias, 'EM_PROGRESSO');
          config.status = 'EM_PROGRESSO';
        }
        
        const tipoProcessamento = config.status === 'ATUALIZADO' 
          ? 'atualização otimizada (últimos 7 dias)' 
          : 'período incremental';
        
        Logger.registrar('INICIO', config.nomeEmpresa, config.alias,
          `Processando ${tipoProcessamento}`, 
          { ultimaData: config.ultimaData });
        
        const resultado = YampiAPI.buscarPedidosIncremental(config);
        
        if (config.status === 'ATUALIZADO') {
          const hoje = new Date();
          ConfigManager.atualizarStatus(config.alias, 'ATUALIZADO', Utilitarios.formatarData(hoje));
          
          Logger.registrar('SUCESSO', config.nomeEmpresa, config.alias,
            'Atualização contínua concluída', 
            { novaData: Utilitarios.formatarData(hoje) });
        } else {
          if (!resultado.completo) {
            Logger.registrar('INFO', config.nomeEmpresa, config.alias,
              'Processamento parcial - aguardando continuacao', {});
            break; // Para e aguarda continuação automática
          }
          
          const dataBase = new Date(config.ultimaData);
          const mesProcessado = dataBase.getMonth() + 1;
          const ultimoDiaMes = new Date(dataBase.getFullYear(), mesProcessado + 1, 0);
          const hoje = new Date();
          const dataFinal = ultimoDiaMes < hoje ? ultimoDiaMes : hoje;
          
          let novoStatus, novaData;
          
          if (Utilitarios.ehMesAtual(dataFinal)) {
            novoStatus = 'ATUALIZADO';
            novaData = Utilitarios.formatarData(hoje);
          } else {
            novoStatus = 'EM_PROGRESSO';
            novaData = Utilitarios.formatarData(resultado.ultimaData);
          }
          
          ConfigManager.atualizarStatus(config.alias, novoStatus, novaData);
          
          Logger.registrar('SUCESSO', config.nomeEmpresa, config.alias,
            'Periodo COMPLETO', { novaData });
        }
        
      } catch (erro) {
        Logger.registrar('ERRO', config.nomeEmpresa, config.alias,
          erro.toString(), { stack: erro.stack });
      }
    }
    
    // Agregação final
    Aggregator.agregarDadosMensais();
    
    // Limpar triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processarProximasEmpresas') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 'Erro ao processar proximas empresas', { erro: erro.toString() });
  }
}

/**
 * ✅ Limpar triggers de continuação automaticamente
 */
function limparTriggersContinuacao() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;
    
    triggers.forEach(trigger => {
      const funcName = trigger.getHandlerFunction();
      const source = trigger.getTriggerSource();
      
      // Remover triggers de continuação e sincronização baseados em tempo
      if (source === ScriptApp.TriggerSource.CLOCK) {
        if (funcName === 'continuarProcessamentoAdaptativo' || 
            funcName === 'executarSincronizacao' ||
            funcName === 'processarProximasEmpresas') {
          try {
            ScriptApp.deleteTrigger(trigger);
            removidos++;
          } catch (e) {
            // Ignorar erros de triggers já removidos
          }
        }
      }
    });
    
    if (removidos > 0) {
      Logger.registrar('INFO', '', '', 
        `${removidos} trigger(s) de continuacao removido(s)`, {});
    }
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 
      'Erro ao limpar triggers de continuacao', { erro: erro.toString() });
  }
}

function obterAliasesPorEmpresa(empresaId) {
  const empresas = ConfigManager.lerConfiguracao();
  const aliases = empresas
    .filter(e => e.abaDestino === empresaId)
    .map(e => e.alias);
  return aliases;
}

/**
 * Retorna a última linha de log para exibir no modal (streaming simples)
 */
function getLatestLogForStreaming() {
  try {
    const sheet = SheetManager.obterAbaLogs();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { message: 'Aguardando primeiros registros de log...' };
    }

    const linha = sheet.getRange(lastRow, 1, 1, 6).getValues()[0];
    const dataHora = linha[0];
    const tipo = linha[1];
    const empresa = linha[2];
    const alias = linha[3];
    const mensagem = linha[4];

    const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
    const dataFmt = dataHora
      ? Utilities.formatDate(new Date(dataHora), tz, 'dd/MM HH:mm')
      : '';

    const prefixo =
      (dataFmt ? `[${dataFmt}] ` : '') +
      (tipo ? `${tipo} - ` : '') +
      [empresa, alias].filter(Boolean).join(' / ');

    return {
      message: `${prefixo} ${mensagem || ''}`.trim()
    };
  } catch (erro) {
    return { message: `Erro ao ler logs: ${erro}` };
  }
}

// ====================================================================//
// TESTES - TEMPORÁRIOS
// ====================================================================//

function limparTodosTriggersTemporarios() {
  const triggers = ScriptApp.getProjectTriggers();
  let contador = 0;
  
  triggers.forEach(trigger => {
    const func = trigger.getHandlerFunction();
    if (func === 'continuarProcessamentoAdaptativo' || 
        func === 'processarProximasEmpresas' ||
        func === 'continuarProcessamento') {
      ScriptApp.deleteTrigger(trigger);
      contador++;
    }
  });
  
  console.log(`${contador} triggers temporários removidos`);
}

function limparTodosEstadosPendentes() {
  const props = PropertiesService.getScriptProperties();
  const allKeys = props.getKeys();
  
  let contador = 0;
  allKeys.forEach(key => {
    if (key.startsWith('ESTADO_')) {
      props.deleteProperty(key);
      contador++;
    }
  });
  
  console.log(`${contador} estados limpos`);
}

/**
 * ✅ Limpeza emergencial de triggers temporários
 * Preserva triggers programados (diários)
 */
function limparTodosTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;
    
    triggers.forEach(trigger => {
      const funcName = trigger.getHandlerFunction();
      const source = trigger.getTriggerSource();
      
      // Remover apenas triggers baseados em tempo, EXCETO os programados
      if (source === ScriptApp.TriggerSource.CLOCK) {
        // ✅ NÃO remover triggers diários importantes
        if (funcName !== 'iniciarSincronizacaoSemanal' && 
            funcName !== 'limparLogsAutomatico') {
          try {
            ScriptApp.deleteTrigger(trigger);
            removidos++;
            Logger.registrar('INFO', '', '', `Trigger removido: ${funcName}`, {});
          } catch (e) {
            // Ignorar
          }
        }
      }
    });
    
    // Limpar propriedades relacionadas
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('INDICE_SYNC');
    props.deleteProperty('PROXIMA_EMPRESA_ABA');
    props.deleteProperty('PROXIMA_EMPRESA_ALIAS');
    
    Logger.registrar('SUCESSO', '', '', 
      `Limpeza emergencial concluida: ${removidos} triggers removidos`, {});
    
    SpreadsheetApp.getUi().alert(
      'Limpeza Concluída',
      `${removidos} trigger(s) temporário(s) removido(s)\n\n` +
      `✅ Triggers diários (02h e 23h) preservados`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (erro) {
    Logger.registrar('ERRO', '', '', 'Erro na limpeza emergencial', 
      { erro: erro.toString() });
  }
}