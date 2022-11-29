const pasta = DriveApp.getFolderById('1JvqEyDaR1o_XKHeGMMz_HsSrSX4ChXG9')

const doc_zerado = DriveApp.getFileById('1iflENCfk8vrDa7VUmgkjN90mUMBPCnNR1ZzOIamuk9s') // Termo Aditivo zerado
const g_doc_zerado = DocumentApp.openById(doc_zerado.getId())

const plan_feitas = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1D8oHoOGh47Tap6nIQQ9XyIcVEOlwL4lB-Dgu4KqO_gU/edit#gid=0")

const total_elements = g_doc_zerado.getBody().getNumChildren()

const plan = SpreadsheetApp.getActiveSpreadsheet()
 
const aba_listas = plan.getSheetByName('LISTAS')
const aba_dados = plan.getSheetByName('v1') //v1
//const aba_dados = plan.getSheetByName('v2') //v2

const dados = aba_dados.getDataRange().getValues()

const regionais = [
  "BARREIRO",
  "CENTRO SUL",
  "LESTE", 
  "NÍVEL CENTRAL",
  "NORDESTE", 
  "NOROESTE", 
  "NORTE", 
  "OESTE", 
  "PAMPULHA", 
  "UNDADE FORA DAS REGIONAIS",
  "VENDA NOVA"
]

function gerarTermo() {

  regionais.forEach(regional => {

    const cabecalho_regionais = aba_listas.getRange("P1:Z2").getValues()

    //while(pastas_regs.hasNext()){
      /*const pasta_reg = pastas_regs.next()*/

      const pasta_reg = pasta.createFolder(regional)
      console.log("Pasta " + regional + " criada")
     // const nome_pasta_reg = pasta_reg.getName()

      //if(nome_pasta_reg === regional){
        cabecalho_regionais[1].forEach((regional_lista,i) => {
          if(regional_lista === regional){
            const n_unidades = cabecalho_regionais[0][i]
            for(let y = 0 ; y < n_unidades; y++){
              
              const unidade = aba_listas.getRange(3+y,16+i).getValue()
              
              console.log("Fazendo a unidade " + unidade)

              const doc_novo = doc_zerado.makeCopy(unidade,pasta_reg)
              const arquivo_id = doc_novo.getId()
              
              const arr_unidade = dados.filter((dado,i)=>{
                if(i=1){
                  return dado[19] === unidade
                }
              })

              arr_unidade.forEach((profissional,i) => {
                  
                const g_doc = DocumentApp.openById(arquivo_id) // v1
              
                const codigo_contrato	= profissional[0]
                const nome = profissional[1]
                const cpf = profissional[2]
                const cargo_efetivo	= profissional[3] === "MEDICO" ? "MÉDICO " : ""
                const especialidade	= profissional[4]
                const data_admissao	= Utilities.formatDate(profissional[5],"GMT","dd/MM/yyyy")
                const conselho_tratado = profissional[7]
                const inscricao_conselho = profissional[8]
                const smsa_regional	= profissional[11]
                const numero_opus	= profissional[12]
                const data_inicio_aditivo	= Utilities.formatDate(profissional[13],"GMT","dd/MM/yyyy")
                const data_fim_aditivo = Utilities.formatDate(profissional[14],"GMT","dd/MM/yyyy")
                const inicio_aditivo_extenso = dataExtenso(profissional[15])
                const pelo_secretario = profissional[16]	
                const secretario = profissional[17]
                const genero = profissional[18]
                const lotacao = profissional[19]

                const corpo = g_doc.getBody()

                corpo.replaceText('{{PROCESSO_OPUS}}',numero_opus)
                corpo.replaceText('{{PELO_SECRETARIO}}',pelo_secretario)
                corpo.replaceText('{{NOME_CONTRATADO}}',nome)
                corpo.replaceText('{{CARGO}}',cargo_efetivo)
                corpo.replaceText('{{ESPECIALIDADE}}',especialidade)
                corpo.replaceText('{{_CPF_}}',cpf)
                corpo.replaceText('{{CONSELHO}}',conselho_tratado)
                corpo.replaceText('{{NUM_CONSELHO}}',inscricao_conselho)
                corpo.replaceText('{{INICIO_ADITIVO}}',data_inicio_aditivo)
                corpo.replaceText('{{FIM_ADITIVO}}',data_fim_aditivo)
                corpo.replaceText('{{INICIO_ADITIVO_EXT}}',inicio_aditivo_extenso)
                corpo.replaceText('{{SECRETARIO}}',secretario)
                corpo.replaceText('{{GENERO}}',genero)
                corpo.replaceText('{{LOTACAO}}',lotacao)
                corpo.replaceText('{{REGIONAL}}',smsa_regional)
                corpo.replaceText('{{DATA_ADMISSAO}}',data_admissao)
                corpo.replaceText('{{MATRICULA}}',codigo_contrato)

                Logger.log(`${i+1}º feito`)

                ultimoParagrafo(g_doc)
                if(i%50 === 0){g_doc.saveAndClose()}
                
              })

              console.log("Unidade " + unidade + " feita")

            }
          }
        })
      //}
    //}

  })
  
}

function dataExtenso(data){
  const dia = data.getDate() > 10 ? data.getDate() : "0" + data.getDate()
  const numero_mes = data.getMonth()
  const ano = data.getFullYear()

  const mes_ext = [
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro"
  ]

  return (dia + " de " + mes_ext[numero_mes] + " de " + ano)

}

function ultimoParagrafo(g_doc){

  //g_doc.getBody().appendParagraph("")

  for(let i = 0; i < total_elements; i++){
    const elemento = g_doc_zerado.getBody().getChild(i).copy()
    const tipo = elemento.getType()
    if(tipo === DocumentApp.ElementType.PARAGRAPH /*&& i !== 29*/){
      g_doc.getBody().appendParagraph(elemento)
    }
    if(tipo === DocumentApp.ElementType.TABLE || tipo === DocumentApp.ElementType.TABLE_CELL || tipo === DocumentApp.ElementType.TABLE_OF_CONTENTS || tipo === DocumentApp.ElementType.TABLE_ROW){
      g_doc.getBody().appendTable(elemento)
      //console.log(i)
    }
    //console.log("elemento " + i + " copiado")
    
  }

}

function compartilharGGTR(){

  const corretos = [
    "ggcat.controladoria@pbh.gov.br", 
    "ggcat.manutencao@pbh.gov.br", 
    "analisecadm4@pbh.gov.br"
  ]

  regionais.forEach((regional,i) => {
    const pastas_regs = pasta.getFolders()
    while(pastas_regs.hasNext()){
      const pasta_reg = pastas_regs.next()
      const nome_pasta_reg = pasta_reg.getName()
      if(nome_pasta_reg === regional){
        console.log(regional)
        const regional_email = aba_listas.getRange(39+i,3).getValue()
        console.log("e-mail da regional: " + regional_email)
        pasta_reg.addEditor(regional_email)
        
        /*pasta_reg.getEditors().forEach(editor => {
          if(!corretos.includes(editor.getEmail())){
            pasta_reg.removeEditor(editor.getEmail())
            console.log("pasta removida")
          }
        })*/

        const arquivos = pasta_reg.getFiles()
        while(arquivos.hasNext()){
          const arquivo = arquivos.next()
          arquivo.addEditor(regional_email)

          /*arquivo.getEditors().forEach(editor => {
            if(!corretos.includes(editor.getEmail())){
              arquivo.removeEditor(editor.getEmail())
              console.log("arquivo removido")
            }
          })*/

          console.log(arquivo.getName())
        }
      }
    }
  })
}

function separarPlanilha(){
  regionais.forEach((regional,i) => {
    
    console.log("Fazendo a regional " + regional)
    
    const dados = aba_dados.getDataRange().getValues()
    const cabecalho = [dados[1]]

    console.log("cabeçalho e dados coletados")

    const arr_regional = dados.filter(dado => {
      return dado[11] === regional
    })

    const linhas = arr_regional.length
    const colunas = arr_regional[0].length

    console.log("Array da regional " + regional + " tem " + linhas + " linhas e " + colunas + " colunas")

    const pastas_regs = pasta.getFolders()
    while(pastas_regs.hasNext()){
      const pasta_reg = pastas_regs.next()
      const nome_pasta_reg = pasta_reg.getName()
      if(nome_pasta_reg === regional){
        
        console.log("Entramos na pasta da regional " + nome_pasta_reg)
        
        const nova_plan = SpreadsheetApp.create("Termos aditivos - " + regional)

        console.log("Planilha criada")

        nova_plan.getSheets()[0].getRange(1,1,1,colunas).setValues(cabecalho)
        nova_plan.getSheets()[0].getRange(2,1,linhas,colunas).setValues(arr_regional)

        console.log("Dados colados")

        const nova_plan_id = nova_plan.getId()
        const arquivos = DriveApp.getFiles()
        while(arquivos.hasNext()){
          const arquivo = arquivos.next()
          const arquivo_id = arquivo.getId()
          if(arquivo_id === nova_plan_id){
            arquivo.moveTo(pasta_reg)
            console.log("Planilha movida para a pasta")
            break
          }
        }
        
      }
    }
  })
}

function corrigirCategoria(){
  
  const profs_sem_categoria = dados.filter((dado => {
    return (dado[3] === "ENFERMEIRO" && dado[4] === "")
  }))

  console.log(profs_sem_categoria.length + " enfermeiros sem categoria")

  profs_sem_categoria.forEach((profissional,i) => {
    console.log(profissional[1])
    const pastas_regs = pasta.getFolders()
    while(pastas_regs.hasNext()){
      const pasta_reg = pastas_regs.next()
      const nome_pasta_reg = pasta_reg.getName()
      if(nome_pasta_reg === profissional[11]){
        console.log("Regional " + profissional[11])
        const arquivos = pasta_reg.getFiles()
        while(arquivos.hasNext()){
          const arquivo = arquivos.next()
          if(arquivo.getName() === profissional[19]){
            console.log("Unidade: " + profissional[19])

            const arquivo_id = arquivo.getId()
            const g_doc = DocumentApp.openById(arquivo_id)

            const corpo = g_doc.getBody()

            corpo.replaceText(profissional[1]+", ,",profissional[1]+", ENFERMEIRO,")

            break

          }
        }
      }
    }
  })
}

