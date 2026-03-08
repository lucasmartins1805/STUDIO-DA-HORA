const URL_API = "https://script.google.com/macros/s/AKfycbyPk9IeHwvyvR04Ep95r0foy8W22el_jSmMm2YKd0NnPxWOKV-vbn3R9EOYRuQajRT1zA/exec"  

//arrumar - esta registrando mesmo sem preencer nada
//excel não esta puxando


async function salvar() {
  // 1. Captura os elementos do HTML
    const campoCliente = document.getElementById('cliente');
    const campoPreco = document.getElementById('preco');
    const campoServico = document.getElementById('servico');

    const dados = {
        cliente: document.getElementById('cliente').value,
        preco: document.getElementById('preco').value,
        servico: document.getElementById('servico').value
    };

    try {
        // Envio 'cego' para evitar o erro de CORS mostrado nas suas imagens
        await fetch(URL_API, {
            method: 'POST',
            mode: 'no-cors', 
            body: JSON.stringify(dados)
        });

        if (campoCliente === "") {
        alert("Digite um nome para consultar.");
        return;
    }
        // 2. O SEGREDO: Limpar logo após o envio bem-sucedido
        alert("Registrado com sucesso!");
        
        campoCliente.value = "";
        campoPreco.value = "";
        campoServico.value = "";
campoCliente.focus();
    } catch (error) {
        console.error("Erro ao salvar:", error);
    }
}

async function buscar() {
    // .trim() remove espaços acidentais no início ou fim
    const nomeBusca = document.getElementById('buscaNome').value.trim().toLowerCase();
    
    if (nomeBusca === "") {
        alert("Digite um nome para consultar.");
        return;
    }

    try {
        const response = await fetch(URL_API);
        const dados = await response.json();

        let visitas = 0;
        let total = 0;
        let achouExato = false;

        dados.forEach(linha => {
            // linha[1] é a coluna do Nome na sua planilha
            const nomeNaPlanilha = linha[1].toString().trim().toLowerCase();

            // COMPARAÇÃO EXATA: só entra se o nome for IDÊNTICO
            if (nomeNaPlanilha === nomeBusca) {
                visitas++;
                total += parseFloat(linha[2]) || 0;
                achouExato = true;
            }
        });

        if (achouExato) {
            // Atualiza o visor conforme seu desenho original
            document.getElementById('vNome').innerText = nomeBusca.toUpperCase();
            document.getElementById('vVisitas').innerText = visitas;
            document.getElementById('vTotal').innerText = total.toFixed(2);
        } else {
            // Se percorreu tudo e não achou o nome exato
            alert("Erro: O nome '" + nomeBusca + "' não foi encontrado exatamente como digitado.");
            limparVisor();
        }

    } catch (error) {
        console.error("Erro na consulta:", error);
        alert("Erro ao conectar com a planilha. Verifique sua internet ou a URL da API.");
    }
}

function limparVisor() {
    document.getElementById('vNome').innerText = "-";
    document.getElementById('vVisitas').innerText = "0";
    document.getElementById('vTotal').innerText = "0.00";
}

function exportarExcel() {
    // 1. Substitua pelo seu ID real aqui
    const planilhaID = "COLOQUE_AQUI_O_ID_DA_SUA_PLANILHA"; 
    
    // 2. Esta URL exporta a aba inteira como XLSX
    const url = "https://docs.google.com/spreadsheets/d/" + planilhaID + "/export?format=xlsx";

    // 3. Tenta abrir o download
    const link = document.createElement('a');
    link.href = url;
    link.target = '_blank';
    link.download = 'Relatorio_Gi_da_Hora.xlsx';
    
    // Simula o clique para contornar bloqueios de pop-up
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}