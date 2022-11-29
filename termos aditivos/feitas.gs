const pasta = DriveApp.getFolderById('1JvqEyDaR1o_XKHeGMMz_HsSrSX4ChXG9')

const doc_zerado = DriveApp.getFileById('1iflENCfk8vrDa7VUmgkjN90mUMBPCnNR1ZzOIamuk9s') // Termo Aditivo zerado
const g_doc_zerado = DocumentApp.openById(doc_zerado.getId())

const plan_feitas = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1D8oHoOGh47Tap6nIQQ9XyIcVEOlwL4lB-Dgu4KqO_gU/edit#gid=0")

const total_elements = g_doc_zerado.getBody().getNumChildren()

const plan = SpreadsheetApp.getActiveSpreadsheet()
 
const aba_listas = plan.getSheetByName('LISTAS')
const aba_dados = plan.getSheetByName('v1') //v1
//const aba_dados = plan.getSheetByName('v2') //v2

const estilo = {}
estilo[DocumentApp.Attribute.FONT_FAMILY] = 'Arial'
estilo[DocumentApp.Attribute.FONT_SIZE] = 9

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

function listarPastas(){
  let linha = 0
  const pastas = pasta.getFolders()
  while(pastas.hasNext()){
    let i = 0
    const pasta = pastas.next()
    const nome_pasta = pasta.getName()
    while(!nome_pasta.includes(regionais[i])){
      i++
    }
    linha++
    plan.getSheetByName("pastas_regionais").getRange(linha,1).setValue(regionais[i])
    plan.getSheetByName("pastas_regionais").getRange(linha,2).setValue(pasta.getUrl())
  }
}

function atualizarNomesPastas(){
  const pastas = pasta.getFolders()
  while(pastas.hasNext()){
    const pasta = pastas.next()
    const nome_pasta = pasta.getName()
    if(regionais.includes(nome_pasta)){
      console.log(nome_pasta)
      pasta.setName("Termos Aditivos - 10/2022 - " + nome_pasta)
    }
  }
  
}

function estilizarDocs() {
  
  regionais.forEach(regional => {
    const pasta_reg = pasta.getFoldersByName(regional).next()
    const docs = pasta_reg.getFiles()

    console.log(regional)

    while(docs.hasNext()){
      const doc = docs.next()
      const doc_id = doc.getId()
      const doc_nome = doc.getName()
      const g_doc = DocumentApp.openById(doc_id)
      const corpo = g_doc.getBody()
      
      //conferindo qual é o texto da última linha de cada página
        
      // setando fonte e tamanho de fonte
        corpo.setAttributes(estilo)
        console.log("Documento " + doc_nome + " estilizado")
    }
  })
  
}
