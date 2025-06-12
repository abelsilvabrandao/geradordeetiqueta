// Atualiza o ano no footer
function atualizarAno() {
  const anoElement = document.getElementById('currentYear');
  const anoAtual = new Date().getFullYear();
  if (anoElement) {
    anoElement.textContent = anoAtual;
  }
}

// Atualiza o ano quando a página carregar
document.addEventListener('DOMContentLoaded', atualizarAno);

let dadosExcel = [];

function excelDateToJSDate(serial) {
  // Excel date serial to JS Date
  if (typeof serial === 'string') {
    // Se já for uma string de data (ex: "09/06/2025"), retorna como está
    return new Date(serial.split('/').reverse().join('-'));
  }

  // Se for número, converte do formato serial do Excel
  const EXCEL_EPOCH = new Date(1899, 11, 30); // Data base do Excel (30/12/1899)
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  
  // Ajusta para o fuso horário local
  const date = new Date(EXCEL_EPOCH.getTime() + (serial - 1) * MS_PER_DAY);
  const localDate = new Date(date.getTime() + date.getTimezoneOffset() * 60 * 1000);
  
  return localDate;
}

document.getElementById("excelFile").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    let json = XLSX.utils.sheet_to_json(sheet, { raw: true });

    // Trim spaces and convert ETA serial numbers to date strings
    json = json.map((row) => {
      // Trim spaces from string fields
      if (row.ETA) {
        if (typeof row.ETA === "string") {
          // Se já for uma string formatada, apenas remove espaços
          row.ETA = row.ETA.trim();
        } else if (typeof row.ETA === "number") {
          // Se for número serial do Excel
          const date = new Date(Math.round((row.ETA - 25569) * 86400 * 1000));
          const day = String(date.getUTCDate()).padStart(2, "0");
          const month = String(date.getUTCMonth() + 1).padStart(2, "0");
          const year = date.getUTCFullYear();
          row.ETA = `${day}/${month}/${year}`;
        }
      }
      
      // Trim other fields
      if (row.Navio__VG_ && typeof row.Navio__VG_ === "string") {
        row.Navio__VG_ = row.Navio__VG_.trim();
      }
      if (row.CLIENTE && typeof row.CLIENTE === "string") {
        row.CLIENTE = row.CLIENTE.trim();
      }
      return row;
    });

    dadosExcel = json;

    // Enable Visualizar and Limpar buttons
    document.getElementById("btnVisualizar").disabled = false;
    document.getElementById("btnLimpar").disabled = false;

    // Clear previous etiquetas and disable Gerar button until visualization
    document.getElementById("etiquetas").innerHTML = "";
    document.getElementById("btnGerar").disabled = true;
  };
  reader.readAsArrayBuffer(file);
});

function renderEtiquetas(dados) {
  const container = document.getElementById("etiquetas");
  container.innerHTML = "";
  const porPagina = 27;

  for (let i = 0; i < dados.length; i += porPagina) {
    const pagina = document.createElement("div");
    pagina.className = "pagina";

    const bloco = dados.slice(i, i + porPagina);
    bloco.forEach((row) => {
      const div = document.createElement("div");
      div.className = "etiqueta";
      div.innerHTML = `
        <div class="etiqueta-line"><span>NAVIO/VG: </span><strong>${row.Navio__VG_ || ""}</strong></div>
        <div class="etiqueta-line"><span>ETA: </span><strong>${row.ETA || ""}</strong></div>
        <div class="etiqueta-line"><span>IMP: </span><strong>${row.CLIENTE || ""}</strong></div>
      `;
      pagina.appendChild(div);
    });

    // Preencher a última página com etiquetas vazias se necessário
    if (bloco.length < porPagina) {
      const faltantes = porPagina - bloco.length;
      for (let j = 0; j < faltantes; j++) {
        const divVazia = document.createElement("div");
        divVazia.className = "etiqueta";
        divVazia.innerHTML = "&nbsp;";
        pagina.appendChild(divVazia);
      }
    }

    container.appendChild(pagina);
  }
}

document.getElementById("btnVisualizar").addEventListener("click", () => {
  if (dadosExcel.length === 0) return;
  renderEtiquetas(dadosExcel);
  document.getElementById("btnGerar").disabled = false;
});

document.getElementById("btnGerar").addEventListener("click", () => {
  const pdf = new jsPDF({
    unit: 'mm',
    format: 'a4',
    orientation: 'portrait'
  });

  // Configurações da página
  const margemX = 10; // 10mm de margem lateral
  const margemY = 10; // 10mm de margem superior/inferior
  const colunas = 3;
  const linhas = 9;
  const larguraUtil = 190; // 210mm - 20mm de margens
  const alturaUtil = 277; // 297mm - 20mm de margens
  const larguraEtiqueta = larguraUtil / colunas;
  const alturaEtiqueta = alturaUtil / linhas;

  // Configurações de fonte
  pdf.setFont('helvetica');
  const fonteTitulo = 8;
  const fonteConteudo = 9;

  function desenharEtiqueta(dados, x, y) {
    // Desenha a borda
    pdf.rect(x, y, larguraEtiqueta - 2, alturaEtiqueta - 2);

    // Configurações de texto
    const margemInternaX = 4;
    const margemInternaY = 4;
    const larguraTextoUtil = larguraEtiqueta - (2 * margemInternaX) - 2;
    let yTexto = y + margemInternaY + 4;

    function quebrarTexto(texto, larguraMaxima) {
      const palavras = texto.split(' ');
      let linhas = [];
      let linhaAtual = '';

      for (let palavra of palavras) {
        const tentativa = linhaAtual + (linhaAtual ? ' ' : '') + palavra;
        if (pdf.getTextWidth(tentativa) <= larguraMaxima) {
          linhaAtual = tentativa;
        } else {
          if (linhaAtual) linhas.push(linhaAtual);
          linhaAtual = palavra;
        }
      }
      if (linhaAtual) linhas.push(linhaAtual);
      return linhas;
    }

    function centralizarLinha(titulo, conteudo) {
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(fonteTitulo);
      const tituloWidth = pdf.getTextWidth(titulo);

      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(fonteConteudo);

      // Calcula espaço disponível para o conteúdo
      const espacoDisponivel = larguraEtiqueta - (2 * margemInternaX) - 2 - tituloWidth;
      const linhas = quebrarTexto(conteudo, espacoDisponivel);

      // Escreve o título
      const xCentro = x + (larguraEtiqueta / 2);
      const xTitulo = xCentro - (tituloWidth + pdf.getTextWidth(linhas[0])) / 2;
      
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(fonteTitulo);
      pdf.text(titulo, xTitulo, yTexto);

      // Escreve a primeira linha do conteúdo
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(fonteConteudo);
      pdf.text(linhas[0], xTitulo + tituloWidth, yTexto);

      // Se houver mais linhas, escreve abaixo
      for (let i = 1; i < linhas.length; i++) {
        yTexto += 4;
        const xLinha = x + (larguraEtiqueta - pdf.getTextWidth(linhas[i])) / 2;
        pdf.text(linhas[i], xLinha, yTexto);
      }

      yTexto += 6; // Espaçamento para próxima informação
    }

    // NAVIO/VG
    centralizarLinha('NAVIO/VG: ', dados.Navio__VG_ || '');

    // ETA
    centralizarLinha('ETA: ', dados.ETA || '');

    // IMP
    centralizarLinha('IMP: ', dados.CLIENTE || '');
  }

  // Gerar páginas
  const porPagina = colunas * linhas;
  let paginaAtual = 1;

  for (let i = 0; i < dadosExcel.length; i++) {
    const posicaoNaPagina = i % porPagina;
    const coluna = posicaoNaPagina % colunas;
    const linha = Math.floor(posicaoNaPagina / colunas);

    const x = margemX + (coluna * larguraEtiqueta);
    const y = margemY + (linha * alturaEtiqueta);

    if (i > 0 && posicaoNaPagina === 0) {
      pdf.addPage();
      paginaAtual++;
    }

    desenharEtiqueta(dadosExcel[i], x, y);
  }

  pdf.save('etiquetas-navio.pdf');
});

document.getElementById("btnLimpar").addEventListener("click", () => {
  dadosExcel = [];
  document.getElementById("excelFile").value = "";
  document.getElementById("etiquetas").innerHTML = "";
  document.getElementById("btnVisualizar").disabled = true;
  document.getElementById("btnGerar").disabled = true;
  document.getElementById("btnLimpar").disabled = true;
});
