/**
 * Trigger executado automaticamente ao editar a planilha.
 * Caso a edição ocorra na aba "PEDIDOS", executa a validação cruzada.
 */
function onEdit(e) {
  const abaEditada = e.source.getActiveSheet().getName();
  if (abaEditada === "PEDIDOS") {
    verificarEAtualizarStatus();
  }
}

/**
 * Normaliza valores numéricos para comparação
 */
function normalizarValor(valor) {
  if (typeof valor === "number") return valor.toString().replace('.', ',');
  if (typeof valor === "string") return valor.trim().replace(/\./g, '').replace(',', '.');
  return valor;
}

/**
 * Converte data no formato brasileiro (dd/mm/yyyy) para objeto Date
 */
function parseDataBrasileira(valor) {
  if (valor instanceof Date) return valor;

  if (typeof valor === "string") {
    const partes = valor.trim().split("/");
    if (partes.length === 3) {
      const dia = parseInt(partes[0], 10);
      const mes = parseInt(partes[1], 10) - 1;
      const ano = parseInt(partes[2], 10);
      return new Date(ano, mes, dia);
    }
  }
  return null;
}

/**
 * Realiza cruzamento entre abas BASE e PEDIDOS
 * Atualiza status ou destaca divergências
 */
function verificarEAtualizarStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBase = ss.getSheetByName("BASE");
  const abaPedidos = ss.getSheetByName("PEDIDOS");

  const dadosBase = abaBase.getDataRange().getValues();
  const dadosPedidos = abaPedidos.getDataRange().getValues();

  for (let i = 1; i < dadosBase.length; i++) {
    const linhaBase = dadosBase[i];

    const codA_Base = String(linhaBase[6]);
    const campoH_Base = String(linhaBase[7]);
    const codProd_Base = String(linhaBase[8]);
    const cliente_Base = String(linhaBase[3]);
    const preco_Base = normalizarValor(linhaBase[13]);
    const qtd_Base = normalizarValor(linhaBase[10]);
    const data_Base = parseDataBrasileira(linhaBase[1]);

    for (let j = 1; j < dadosPedidos.length; j++) {
      const linhaPed = dadosPedidos[j];

      const codA_Ped = String(linhaPed[0]);
      const campoH_Ped = String(linhaPed[1]);
      const codProd_Ped = String(linhaPed[2]);
      const cliente_Ped = String(linhaPed[3]);
      const preco_Ped = normalizarValor(linhaPed[4]);
      const qtd_Ped = normalizarValor(linhaPed[5]);
      const data_Ped = parseDataBrasileira(linhaPed[6]);

      if (!data_Base || !data_Ped) continue;

      const codA_OK = codA_Base == codA_Ped;
      const campoH_OK = campoH_Base == campoH_Ped;
      const cliente_OK = cliente_Base == cliente_Ped;
      const preco_OK = preco_Base == preco_Ped;
      const data_OK = data_Base < data_Ped;

      const restanteOK = codA_OK && campoH_OK && cliente_OK && preco_OK && data_OK;

      const codProd_OK = codProd_Base == codProd_Ped;
      const qtd_OK = qtd_Base == qtd_Ped;

      if (restanteOK) {
        if (codProd_OK && qtd_OK) {
          abaBase.getRange(i + 1, 1).setValue("ENCERRADO");
        } else if (!codProd_OK && qtd_OK) {
          abaBase.getRange(i + 1, 9).setBackground("#FFF59D");
        } else if (codProd_OK && !qtd_OK) {
          abaBase.getRange(i + 1, 11).setBackground("#FFF59D");
        } else {
          abaBase.getRange(i + 1, 9).setBackground("#FFF59D");
          abaBase.getRange(i + 1, 11).setBackground("#FFF59D");
        }
        break;
      }
    }
  }
}
