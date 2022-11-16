const plan = SpreadsheetApp.getActiveSpreadsheet()
const abas = plan.getSheets()
const aba_impressao = plan.getSheetByName("Impressão crescer")
const ui = SpreadsheetApp.getUi()

function onOpen() {

  ui.createMenu('Folha de Ponto')
    .addItem('Novo mês', 'criarNovoMesEAtualizar')
    .addToUi();

  atualizarImpressao()
}

function atualizarImpressao(){
  
  const nomes_abas = abas.map(aba => {
    
    if(aba.getName() !== "Impressão crescer"){
      return aba.getName()
    }
    
  })

  const regra = SpreadsheetApp.newDataValidation().requireValueInList(nomes_abas).build()

  aba_impressao.getRange("k1").setDataValidation(regra)
  console.log("impressao atualiada")
}

function criarNovoMes(){
  const meses = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro"
  ]

  const num_meses = abas.map(aba => {
    const nome = aba.getName()
    if(nome !== "Impressão crescer"){
      return aba.getRange("B1").getValue()
    }
  }).filter(num => num !== undefined)

  const maior_num_mes = Math.max.apply(null,num_meses)

  if(maior_num_mes < 12){
    const nova_aba = plan.getSheets()[0].copyTo(plan)
    nova_aba.setName(meses[maior_num_mes])
    nova_aba.getRange("B4:e34").clearContent()
    nova_aba.getRange("h4:i34").clearContent()
    nova_aba.getRange("b1").setValue(maior_num_mes+1)

    plan.setActiveSheet(nova_aba)
    plan.moveActiveSheet(1)

    return true

  }else{
    ui.alert("Criar uma cópia da planilha atual.")
    return false
  }
}

async function criarNovoMesEAtualizar(){
  await criarNovoMes()
  atualizarImpressao()
}