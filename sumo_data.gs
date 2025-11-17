/**
 * Arquivo: sumo_api_integration.gs
 * Descrição: Script Google Apps Script para integrar a planilha com a API do Sumo.
 * Autor: Manus AI
 */

// A constante BASHO_ID será agora definida dinamicamente dentro da função updateSumoResults.
const DIVISION = "Makuuchi"; // Divisão (Ex: Makuuchi)
const RIKISHI_COLUMN_INDEX = 1; // Coluna A (índice 1)
const PEPPERONI_COLUMN = 2; // Coluna B (índice 2) - Para marcar com "S"
const BASHO_LIST_COLUMN = 3; // Coluna C (índice 3)
const TEMPLATE_SHEET_NAME = "11/2025"; // Nome da aba modelo a ser copiada
const FIRST_DATA_ROW = 2; // Primeira linha de dados (A2)
const WIN_RATE_LABEL = "Win Rate (W/Total)"; // Rótulo da última linha a ser preservada

/**
 * Função auxiliar para obter e validar o ID do Basho a partir do nome da aba.
 * O nome da aba deve estar no formato "mm/yyyy".
 * @returns {string|null} O ID do Basho no formato "yyyyMM" ou null em caso de erro.
 */
function getBashoIdFromSheetName() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  
  // Expressão regular para validar o formato mm/yyyy
  const regex = /^(\d{2})\/(\d{4})$/;
  const match = sheetName.match(regex);
  
  if (!match) {
    ui.alert(
      'Erro de Formato da Aba',
      `O nome da aba ativa ("${sheetName}") não está no formato esperado "mm/yyyy". Por favor, renomeie a aba.`,
      ui.ButtonSet.OK
    );
    return null;
  }
  
  const month = match[1]; // mm
  const year = match[2];  // yyyy
  
  // Validação básica de mês (1 a 12)
  const monthInt = parseInt(month, 10);
  if (monthInt < 1 || monthInt > 12) {
    ui.alert(
      'Erro de Validação de Data',
      `O mês ("${month}") no nome da aba é inválido. O mês deve ser um número entre 01 e 12.`,
      ui.ButtonSet.OK
    );
    return null;
  }
  
  // Converte para o formato yyyyMM (Ex: 202511)
  return `${year}${month}`;
}

/**
 * Função auxiliar para solicitar e validar o ID do Basho ao usuário.
 * @returns {object|null} Um objeto com {bashoId: "yyyyMM", bashoDisplay: "mm/yyyy"} ou null em caso de erro/cancelamento.
 */
function promptAndValidateBashoId() {
  const ui = SpreadsheetApp.getUi();
  
  const bashoResponse = ui.prompt(
    'Basho ID',
    'Por favor, insira o ID do Basho no formato mm/yyyy (Ex: 11/2025):',
    ui.ButtonSet.OK_CANCEL
  );

  if (bashoResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Operação cancelada pelo usuário.');
    return null;
  }

  const input = bashoResponse.getResponseText().trim();
  const regex = /^(\d{2})\/(\d{4})$/;
  const match = input.match(regex);

  if (!match) {
    ui.alert('Erro: O formato inserido não é válido. Use mm/yyyy (Ex: 11/2025).');
    return null;
  }

  const month = match[1];
  const year = match[2];
  
  const monthInt = parseInt(month, 10);
  if (monthInt < 1 || monthInt > 12) {
    ui.alert('Erro: O mês inserido é inválido. O mês deve ser um número entre 01 e 12.');
    return null;
  }
  
  // Retorna no formato yyyyMM e o formato de exibição mm/yyyy
  return {
    bashoId: `${year}${month}`,
    bashoDisplay: input
  };
}

/**
 * Função principal que é executada ao clicar no botão.
 * Solicita o dia, busca os dados da API e atualiza a planilha.
 */
function updateSumoResults() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Obter o ID do Basho a partir do nome da aba
  const BASHO_ID = getBashoIdFromSheetName();
  if (!BASHO_ID) {
    return; // Interrompe a execução se o formato da aba for inválido
  }
  
  // 2. Solicitar o dia ao usuário
  const dayResponse = ui.prompt(
    'Atualizar Resultados do Sumo',
    'Por favor, insira o número do dia (1 a 15) para o qual deseja buscar os resultados:',
    ui.ButtonSet.OK_CANCEL
  );

  // Verificar se o usuário cancelou
  if (dayResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Operação cancelada pelo usuário.');
    return;
  }

  const day = parseInt(dayResponse.getResponseText());

  // Validar o dia
  if (isNaN(day) || day < 1 || day > 15) {
    ui.alert('Erro: O dia inserido deve ser um número entre 1 e 15.');
    return;
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // 3. Ler a lista de Rikishi (Competidores) da Coluna A
    // 3. Ler a lista de Rikishi (Competidores) da Coluna A
    // Encontrar a linha do Win Rate dinamicamente
    const winRateRowIndex = values.findIndex(row => String(row[RIKISHI_COLUMN_INDEX - 1]).trim() === WIN_RATE_LABEL);
    
    // Se não encontrar o rótulo, assume-se que a última linha é o Win Rate (comportamento original)
    const lastRikishiRowIndex = winRateRowIndex !== -1 ? winRateRowIndex : values.length - 1;

    // Ignora o cabeçalho (primeira linha) e a linha do Win Rate (última linha)
    const rikishiList = values.slice(1, lastRikishiRowIndex)
      .map(row => row[RIKISHI_COLUMN_INDEX - 1]) // Coluna A é o índice 0
      .filter(name => name && String(name).trim() !== '') // Filtra nomes vazios
      .map(name => String(name).trim());

    if (rikishiList.length === 0) {
      ui.alert('Erro: Nenhuma lista de competidores encontrada na Coluna A.');
      return;
    }
    
    // 4. Chamar a API do Sumo
    const apiUrl = `https://sumo-api.com/api/basho/${BASHO_ID}/torikumi/${DIVISION}/${day}`;
    const response = UrlFetchApp.fetch(apiUrl);
    const torikumiData = JSON.parse(response.getContentText());
    
    // 5. Processar os dados e atualizar a planilha
    processAndWriteResults(sheet, rikishiList, torikumiData, day);

    ui.alert(`Resultados do Dia ${day} (Basho ID: ${BASHO_ID}) atualizados com sucesso!`);

  } catch (e) {
    ui.alert('Ocorreu um erro durante a execução: ' + e.toString());
  }
}

/**
 * Processa os dados da API e escreve os resultados na planilha.
 */
function processAndWriteResults(sheet, rikishiList, torikumiData, day) {
  const ui = SpreadsheetApp.getUi();
  const headerRow = 1; // Linha do cabeçalho
  
  // 1. Encontrar a coluna de destino
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  
  let targetCol = lastCol + 1;
  let targetColLetter = String.fromCharCode(64 + targetCol); // Converte índice para letra (A=1, B=2, etc.)
  
  // Verificar se a coluna para o dia já existe
  const dayHeader = `Dia ${day}`;
  const existingColIndex = headers.indexOf(dayHeader);
  
  if (existingColIndex !== -1) {
    targetCol = existingColIndex + 1; // Coluna existente (índice + 1)
    targetColLetter = String.fromCharCode(64 + targetCol);
    ui.alert(`A coluna "${dayHeader}" já existe. Os dados serão sobrescritos na coluna ${targetColLetter}.`);
  } else {
    // Adicionar novo cabeçalho
    sheet.getRange(headerRow, targetCol).setValue(dayHeader);
    // Aplicar formatação básica ao cabeçalho
    sheet.getRange(headerRow, targetCol).setFontWeight('bold').setBackground('#cccccc').setBorder(true, true, true, true, true, true);
  }

  // 2. Mapear Rikishi para a linha na planilha
  const rikishiToRowMap = {};
  // A lista de Rikishi começa na FIRST_DATA_ROW e vai até a penúltima linha
  const rikishiNamesRange = sheet.getRange(FIRST_DATA_ROW, RIKISHI_COLUMN_INDEX, rikishiList.length, 1);
  const rikishiNames = rikishiNamesRange.getValues().flat();
  
  rikishiNames.forEach((name, index) => {
    if (name && String(name).trim() !== '') {
      rikishiToRowMap[String(name).trim()] = FIRST_DATA_ROW + index;
    }
  });

  // 3. Processar os resultados da API
  const results = {}; // {Rikishi: 'W' | 'L' | 'VS'}
  const directMatches = []; // [{rikishi1: 'Nome1', rikishi2: 'Nome2', row1: 2, row2: 5}]
  
  // Criar um conjunto para busca rápida
  const rikishiSet = new Set(rikishiList);

  for (const match of torikumiData.torikumi) {
    const eastRikishi = match.eastShikona;
    const westRikishi = match.westShikona;
    const winnerRikishi = match.winnerEn;
    
    const eastInList = rikishiSet.has(eastRikishi);
    const westInList = rikishiSet.has(westRikishi);

    if (eastInList && westInList) {
      // Confronto direto entre dois competidores da lista
      directMatches.push({
        rikishi1: eastRikishi,
        rikishi2: westRikishi,
        row1: rikishiToRowMap[eastRikishi],
        row2: rikishiToRowMap[westRikishi]
      });
      results[eastRikishi] = 'VS';
      results[westRikishi] = 'VS';
      
    } else if (eastInList) {
      // Rikishi da lista lutou contra um de fora
      results[eastRikishi] = (eastRikishi === winnerRikishi) ? 'W' : 'L';
      
    } else if (westInList) {
      // Rikishi da lista lutou contra um de fora
      results[westRikishi] = (westRikishi === winnerRikishi) ? 'W' : 'L';
    }
  }

  // 4. Escrever os resultados na planilha e aplicar formatação
  let totalWins = 0;
  let totalLosses = 0;
  let totalMatches = 0;
  
  const resultsToWrite = [];
  
  // Preencher a coluna de resultados
  for (let i = 0; i < rikishiList.length; i++) {
    const rikishi = rikishiList[i];
    const result = results[rikishi] || ''; // Pode ser vazio se o rikishi não lutou
    resultsToWrite.push([result]);
    
    if (result === 'W') {
      totalWins++;
      totalMatches++;
    } else if (result === 'L') {
      totalLosses++;
      totalMatches++;
    } else if (result === 'VS') {
      // Confrontos diretos não contam para o total de W/L do dia, pois o resultado não é W ou L.
    }
  }
  
  // Escrever os resultados
  const dataRange = sheet.getRange(FIRST_DATA_ROW, targetCol, rikishiList.length, 1);
  dataRange.setValues(resultsToWrite);
  
  // Aplicar formatação padrão aos dados
  dataRange.setBorder(true, true, true, true, true, true);

  // Lista de cores claras pré-definidas (garante que sejam claras)
  let lightColors = ['#FFC7CE', '#B4C6E7', '#C6E0B4', '#FFE699', '#F8CBAD', '#D9D2E9', '#F4CCCC', '#E6B8AF', '#B4A7D6', '#A2C4C9', '#B6D7A8', '#F9CB9C'];

  // Embaralhar a lista para aleatoriedade
  lightColors = shuffleArray(lightColors);

  // Aplicar formatação condicional para confrontos diretos
  for (let i = 0; i < directMatches.length; i++) {
    const match = directMatches[i];
    const color = lightColors[i % lightColors.length];
    
    // Formatar as células dos dois competidores envolvidos
    sheet.getRange(match.row1, targetCol).setBackground(color);
    sheet.getRange(match.row2, targetCol).setBackground(color);
  }
  
  // 5. Calcular e escrever o Win Rate e Totais de W/L
  // A linha do Win Rate é a linha logo após o último Rikishi, que é a linha onde o Win Rate deveria estar.
  const winRateRow = FIRST_DATA_ROW + rikishiList.length;
  
  // Linha Única: Win Rate (W/Total (XX.XX%))
  const winRatePercentage = totalMatches > 0 ? (totalWins / totalMatches) * 100 : 0;
  const winRateDisplay = totalMatches > 0 ? `${totalWins}/${totalMatches} (${winRatePercentage.toFixed(2)}%)` : 'N/A';
  
  const winRateCell = sheet.getRange(winRateRow, targetCol);
  winRateCell.setValue(winRateDisplay)
             .setFontWeight('bold')
             .setBorder(true, true, true, true, true, true);

  /**
   * Função auxiliar para embaralhar um array (Fisher-Yates shuffle).
   * @param {Array<any>} array O array a ser embaralhado.
   * @returns {Array<any>} O array embaralhado.
   */
  function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
  }

}

/**
 * FUNÇÃO 2: Busca a lista de Rikishi (Banzuke) para um Basho específico
 * e atualiza a Coluna A da aba "Pepperoni", adicionando apenas os novos nomes,
 * e atualiza a Coluna C com a lista de Bashos que o Rikishi participou.
 */
function updateRikishiList() {
  const ui = SpreadsheetApp.getUi();
  const SHEET_NAME = 'Pepperoni';
  const TARGET_COLUMN = RIKISHI_COLUMN_INDEX; // Coluna A
  const BASHO_COL = BASHO_LIST_COLUMN; // Coluna C
  
  // 1. Solicitar e validar o BASHO_ID do usuário
  const bashoInfo = promptAndValidateBashoId();
  if (!bashoInfo) {
    return; // Interrompe a execução se o formato for inválido ou cancelado
  }
  
  const BASHO_ID = bashoInfo.bashoId;
  const BASHO_DISPLAY = bashoInfo.bashoDisplay; // Formato mm/yyyy
  
  try {
    // 2. Acessar a aba "Pepperoni"
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      ui.alert(`Erro: A aba "${SHEET_NAME}" não foi encontrada. Por favor, crie a aba.`);
      return;
    }
    
    // 3. Chamar a API do Sumo para o Banzuke
    const apiUrl = `https://sumo-api.com/api/basho/${BASHO_ID}/banzuke/${DIVISION}`;
    const response = UrlFetchApp.fetch(apiUrl);
    const banzukeData = JSON.parse(response.getContentText());
    
    // 4. Extrair a lista de Rikishi (shikonaEn)
    let newRikishiNames = [];
    
    // Extrai Rikishi do lado East
    if (banzukeData.east) {
      newRikishiNames = newRikishiNames.concat(banzukeData.east.map(item => item.shikonaEn).filter(name => name && name.trim() !== ''));
    }
    
    // Extrai Rikishi do lado West
    if (banzukeData.west) {
      newRikishiNames = newRikishiNames.concat(banzukeData.west.map(item => item.shikonaEn).filter(name => name && name.trim() !== ''));
    }
    
    // Remove duplicatas
    newRikishiNames = [...new Set(newRikishiNames)];
    
    if (newRikishiNames.length === 0) {
      ui.alert(`Aviso: Nenhuma lista de Rikishi encontrada para o Basho ID ${BASHO_ID}.`);
      return;
    }
    
    // 5. Ler a lista atual de Rikishi na planilha (Coluna A) e a lista de Bashos (Coluna C)
    const lastRow = sheet.getLastRow();
    let existingRikishi = [];
    let existingBashos = [];
    
    // Verifica se há dados a partir da linha 2
    if (lastRow >= FIRST_DATA_ROW) {
      const numRows = lastRow - FIRST_DATA_ROW + 1;
      
      // Lê Coluna A e Coluna C em uma única chamada para otimizar
      const dataRange = sheet.getRange(FIRST_DATA_ROW, TARGET_COLUMN, numRows, BASHO_COL);
      const values = dataRange.getValues();
      
      let lastFilledRow = FIRST_DATA_ROW - 1; // Começa antes da primeira linha de dados
      
      // Percorre de baixo para cima para encontrar a última célula não vazia na Coluna A
      for (let i = values.length - 1; i >= 0; i--) {
        if (String(values[i][0]).trim() !== '') {
          lastFilledRow = FIRST_DATA_ROW + i;
          break;
        }
      }
      
      // Se encontrou dados, lê apenas até a última linha preenchida
      if (lastFilledRow >= FIRST_DATA_ROW) {
        const finalNumRows = lastFilledRow - FIRST_DATA_ROW + 1;
        // Lê o intervalo final
        const finalValues = sheet.getRange(FIRST_DATA_ROW, TARGET_COLUMN, finalNumRows, BASHO_COL).getValues();
        
        // CORREÇÃO: Garante que os valores lidos da Coluna C sejam tratados como strings.
        existingRikishi = finalValues.map(row => String(row[0]).trim()); // Coluna A
        existingBashos = finalValues.map(row => {
          const bashoValue = row[BASHO_COL - TARGET_COLUMN];
          // Se for um objeto Date (o que causa o problema), formata para string "mm/yyyy"
          if (bashoValue instanceof Date) {
            // Cria o formato mm/yyyy a partir do objeto Date
            const month = (bashoValue.getMonth() + 1).toString().padStart(2, '0');
            const year = bashoValue.getFullYear().toString();
            return `${month}/${year}`;
          }
          return String(bashoValue).trim();
        });
      }
    }
    
    // Mapeia Rikishi existente para sua linha e lista de bashos
    const rikishiMap = {}; // { RikishiName: { row: 2, bashos: "01/2025,03/2025" } }
    existingRikishi.forEach((name, index) => {
      if (name) {
        rikishiMap[name] = {
          row: FIRST_DATA_ROW + index,
          bashos: existingBashos[index]
        };
      }
    });
    
    // 6. Processar Rikishi e preparar dados para escrita
    const rikishiToAdd = [];
    const bashosToUpdate = []; // Array para escrita na Coluna C
    let updateCount = 0;
    
    for (const name of newRikishiNames) {
      const trimmedName = name.trim();
      
      if (rikishiMap[trimmedName]) {
        // Rikishi JÁ EXISTE: Atualizar Coluna C
        const currentBashos = rikishiMap[trimmedName].bashos;
        const bashoArray = currentBashos.split(',').map(b => b.trim());
        
        if (!bashoArray.includes(BASHO_DISPLAY)) {
          // Adiciona o novo Basho se ainda não estiver na lista
          const newBashos = currentBashos ? `${currentBashos},${BASHO_DISPLAY}` : BASHO_DISPLAY;
          bashosToUpdate.push({
            row: rikishiMap[trimmedName].row,
            value: newBashos
          });
          updateCount++;
        }
        
      } else {
        // Rikishi NÃO EXISTE: Adicionar à lista de novos
        rikishiToAdd.push([trimmedName, '', BASHO_DISPLAY]); // Coluna A, B (vazia), C (novo basho)
      }
    }
    
    // 7. Escrever os novos Rikishi (Coluna A e C)
    if (rikishiToAdd.length > 0) {
      // Determina a linha de início para os novos Rikishi
      const lastFilledRow = existingRikishi.length > 0 ? FIRST_DATA_ROW + existingRikishi.length - 1 : FIRST_DATA_ROW - 1;
      const startRow = lastFilledRow < FIRST_DATA_ROW ? FIRST_DATA_ROW : lastFilledRow + 1;
      
      // Escreve os novos Rikishi nas Colunas A, B e C
      const newRikishiRange = sheet.getRange(startRow, TARGET_COLUMN, rikishiToAdd.length, BASHO_COL);
      newRikishiRange.setValues(rikishiToAdd);
      // Adiciona bordas aos novos rikishis (Solicitação do usuário 1)
      newRikishiRange.setBorder(true, true, true, true, true, true);
    }
    
    // 8. Atualizar a Coluna C para os Rikishi existentes
    if (bashosToUpdate.length > 0) {
      // Agrupa as atualizações por linha para otimizar a escrita (melhor performance)
      const rowsToUpdate = bashosToUpdate.map(item => item.row);
      const minRow = Math.min(...rowsToUpdate);
      const maxRow = Math.max(...rowsToUpdate);
      const numRows = maxRow - minRow + 1;
      
      // Lê o intervalo completo da Coluna C
      const rangeToRead = sheet.getRange(minRow, BASHO_COL, numRows, 1);
      const currentValues = rangeToRead.getValues();
      
      // Atualiza o array de valores
      bashosToUpdate.forEach(item => {
        const rowIndex = item.row - minRow;
        currentValues[rowIndex][0] = item.value;
      });
      
      // Escreve o intervalo atualizado de volta na Coluna C
      rangeToRead.setValues(currentValues);
    }
    
    const totalChanges = rikishiToAdd.length + updateCount;
    
    if (totalChanges > 0) {
      ui.alert(`${rikishiToAdd.length} novos Rikishi adicionados e ${updateCount} listas de Basho atualizadas para o Basho ID ${BASHO_DISPLAY} na aba "${SHEET_NAME}"!`);
    } else {
      ui.alert(`Aviso: Nenhum Rikishi novo ou atualização de Basho necessária para o Basho ID ${BASHO_DISPLAY}.`);
    }
    
  } catch (e) {
    ui.alert('Ocorreu um erro durante a execução: ' + e.toString());
  }
}

/**
 * FUNÇÃO 3: Cria uma nova aba copiando a estrutura de uma aba modelo
 * e a renomeia com o BashoId fornecido pelo usuário.
 */
function createNewBashoSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const PEPPERONI_SHEET_NAME = 'Pepperoni';
  
  // 1. Solicitar e validar o nome da nova aba (Basho ID)
  const bashoInfo = promptAndValidateBashoId();
  if (!bashoInfo) {
    return;
  }
  
  const newSheetName = bashoInfo.bashoDisplay;
  
  // 2. Verificar se a aba modelo existe
  const templateSheet = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  if (!templateSheet) {
    ui.alert(`Erro: A aba modelo "${TEMPLATE_SHEET_NAME}" não foi encontrada. Por favor, crie uma aba com este nome para usar como modelo.`);
    return;
  }
  
  // 3. Verificar se a aba de destino já existe
  if (ss.getSheetByName(newSheetName)) {
    ui.alert(`Erro: A aba "${newSheetName}" já existe. Por favor, escolha outro nome ou use a aba existente.`);
    return;
  }
  
  // 4. Acessar a aba "Pepperoni" para obter a lista de Rikishi filtrada
  const pepperoniSheet = ss.getSheetByName(PEPPERONI_SHEET_NAME);
  if (!pepperoniSheet) {
    ui.alert(`Erro: A aba "${PEPPERONI_SHEET_NAME}" não foi encontrada. É necessária para filtrar os Rikishi.`);
    return;
  }
  
  try {
    // 5. Copiar a aba modelo
    const newSheet = templateSheet.copyTo(ss);
    
    // 6. Renomear a nova aba
    newSheet.setName(newSheetName);
    
    // 7. Obter a lista de Rikishi da aba "Pepperoni"
    const pepperoniLastRow = pepperoniSheet.getLastRow();
    let filteredRikishi = [];
    
    if (pepperoniLastRow >= FIRST_DATA_ROW) {
      // Lê Colunas A, B e C da aba "Pepperoni"
      const pepperoniData = pepperoniSheet.getRange(FIRST_DATA_ROW, RIKISHI_COLUMN_INDEX, pepperoniLastRow - FIRST_DATA_ROW + 1, BASHO_LIST_COLUMN).getValues();
      
      // Filtra os Rikishi
      filteredRikishi = pepperoniData
        .filter(row => {
          const pepperoniMark = String(row[PEPPERONI_COLUMN - 1]).trim(); // Coluna B
          const bashoList = String(row[BASHO_LIST_COLUMN - 1]).trim(); // Coluna C
          
          // Verifica se tem "S" na Coluna B E se o Basho ID está na Coluna C
          return pepperoniMark === 'S' && bashoList.split(',').map(b => b.trim()).includes(newSheetName);
        })
        .map(row => [row[RIKISHI_COLUMN_INDEX - 1]]); // Pega apenas o nome do Rikishi (Coluna A)
    }
    
    // 8. Limpar o conteúdo da Coluna A da nova aba (para remover a lista do modelo)
    const newSheetLastRow = newSheet.getLastRow();
    
    // 8a. Limpa o conteúdo da Coluna A, da Linha 2 até a última linha
    if (newSheetLastRow >= FIRST_DATA_ROW) {
      newSheet.getRange(FIRST_DATA_ROW, RIKISHI_COLUMN_INDEX, newSheetLastRow - FIRST_DATA_ROW + 1, 1).clearContent();
    }
    

    
    // 9. Escrever a lista de Rikishi filtrada na Coluna A da nova aba
    if (filteredRikishi.length > 0) {
      // Insere os Rikishi a partir da Linha 2
      newSheet.getRange(FIRST_DATA_ROW, RIKISHI_COLUMN_INDEX, filteredRikishi.length, 1).setValues(filteredRikishi);
      
      // 9b. Adiciona a linha "Win Rate" logo abaixo do último Rikishi
      const winRateRow = FIRST_DATA_ROW + filteredRikishi.length;
      newSheet.getRange(winRateRow, RIKISHI_COLUMN_INDEX).setValue(WIN_RATE_LABEL).setFontWeight('bold');
      
      // 9c. Remove linhas extras (se a lista filtrada for menor que a lista do modelo)
      if (newSheetLastRow > winRateRow) {
        newSheet.deleteRows(winRateRow + 1, newSheetLastRow - winRateRow);
      }
      
    } else {
      // Se não houver Rikishi, apenas adiciona o cabeçalho e a linha Win Rate
      newSheet.getRange(FIRST_DATA_ROW, RIKISHI_COLUMN_INDEX).setValue(WIN_RATE_LABEL).setFontWeight('bold');
      
      // Remove linhas extras
      if (newSheetLastRow > FIRST_DATA_ROW) {
        newSheet.deleteRows(FIRST_DATA_ROW + 1, newSheetLastRow - FIRST_DATA_ROW);
      }
      
      ui.alert(`Aviso: Nenhuma Rikishi encontrado na aba "Pepperoni" com marca "S" e Basho "${newSheetName}". A nova aba foi criada, mas a lista de Rikishi está vazia.`);
    }
    
    // 10. Limpar as colunas de resultados (a partir da Coluna B)
    const lastCol = newSheet.getLastColumn();
    
    if (lastCol > RIKISHI_COLUMN_INDEX) {
      // Remove todas as colunas a partir da Coluna B
      newSheet.deleteColumns(RIKISHI_COLUMN_INDEX + 1, lastCol - RIKISHI_COLUMN_INDEX);
    }
    
    // CORREÇÃO: Copia o conteúdo das linhas A24 e A25 da aba modelo
    // Isso é feito por último para garantir que não seja sobrescrito pela lista de Rikishi.
    const rangeToCopy = templateSheet.getRange('A24:A25');
    const destinationRange = newSheet.getRange('A24:A25');
    rangeToCopy.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    
    ui.alert(`Nova aba "${newSheetName}" criada com sucesso com ${filteredRikishi.length} Rikishi filtrados!`);
    
  } catch (e) {
    ui.alert('Ocorreu um erro durante a criação da aba: ' + e.toString());
  }
}