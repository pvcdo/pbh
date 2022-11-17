function carregarBms(){

  '(?i)(\W|^)(baloney|darn|drat|fooey|gosh\sdarnit|heck)(\W|$)'
  
  var patternBM = "^(" + bms.map(function([e]) {return e}).join("|") + ")$";
  var bmsVal = FormApp.createTextValidation()
    .setHelpText("BM não encontrado no sistema. Certifique-se de estar digitando o BM sem o hífen.")
    .requireTextMatchesPattern(patternBM).build();
  inputMatricula.setValidation(bmsVal)
 
  var patternCPF = "^(" + cpfs.map(function([e]) {return e}).join("|") + ")$";
  var cpfsVal = FormApp.createTextValidation()
    .setHelpText("CPF não encontrado no sistema. Certifique-se de estar digitando apenas números.")
    .requireTextMatchesPattern(patternCPF).build();
  inputCPF.setValidation(cpfsVal)

}