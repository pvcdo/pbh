const plan_admp = SpreadsheetApp.openById("1KGBMHaP6stDkRwh0wmUpr0RCasbvSiIP0uBZ0v9Jv7I").getSheetByName("ADMP")
const form = FormApp.getActiveForm()
const inputMatricula = form.getItems()[1].asTextItem()
const inputCPF = form.getItems()[2].asTextItem()
const bms = plan_admp.getRange("D3:D").getValues() 
const cpfs = plan_admp.getRange("B3:B").getValues()

function itens(){
  form.getItems().forEach((item,i) => {
    if(item.getTitle() === "E-MAIL DO GESTOR"){
      console.log(item.getTitle() + ", id: " + i)
    }
    
  })
}

function formBuilder(vals){
  const formComDados = FormApp.create('Form de ' + vals[5])
  return formComDados
}

function formFiller(formComDados,vals){

  var item = form.addMultipleChoiceItem();

  var sec_dadosProfissional = form.addPageBreakItem()
      .setTitle('Dados do profissional')
      item.setChoices([
        item.createChoice(vals[0],)
      ])
  var page3 = form.addPageBreakItem()
      .setTitle('Third page');

  item.setTitle('Question')
      .setChoices([
         item.createChoice('Yes',page2),
         item.createChoice('No',page3)
       ]);
       
  formComDados.addPageBreakItem().setTitle(vals[0])
  
  formComDados.addPageBreakItem().setTitle('Motivo da rescisão')
  form.addMultipleChoiceItem().setChoices()
  form.addMultipleChoiceItem().createChoice('Por interesse do profissional')

  
}