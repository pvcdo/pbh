const esta_plan = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
const aba_admp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admp")
const aba_class = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("class")

async function executar(){
  console.log("Iniciando atualização da planilha")
  await atualizar()
  await removerLinhas()
  console.log("Iniciando padronização da planilha")
  await padronizar()
  console.log("Iniciando classificação das lotações")
  await passarClass()
  console.log("Iniciando população das especialidades")
  await popularEspecialidades()
}

function atualizar() {
  return new Promise(resolve => {
    const arquivos = DriveApp.getFolderById("1TZs0_1xf7Gbm-YFmcvDnvCb76xP_b0nc").getFiles()
    while(arquivos.hasNext()){
      const arquivo = arquivos.next()
      const arquivo_nome = arquivo.getName()
      if(!arquivo_nome.split("_")[1]){
        const blob = arquivo.getBlob()

        let config = {
          title: arquivo_nome,
          parents: [{id: "1TZs0_1xf7Gbm-YFmcvDnvCb76xP_b0nc"}],
          mimeType: MimeType.GOOGLE_SHEETS
        };
        let plan_atualizada = Drive.Files.insert(config,blob)
        plan_atualizada = SpreadsheetApp.openById(plan_atualizada.id)

        console.log("Planilha cópia criada")
        
        const dados_atualizados = plan_atualizada.getDataRange().getValues()

        const x_dados_atualizados = dados_atualizados[0].length
        const y_dados_atualizados = dados_atualizados.length

        esta_plan.getDataRange().clear()
        console.log("Dados desta planilha apagados")

        esta_plan.getRange(1,1,y_dados_atualizados,x_dados_atualizados).setValues(dados_atualizados)

        console.log("Dados copiados da cópia para a monitoramento")

        //está excluindo alguma outra coisa
        //DriveApp.removeFile(DriveApp.getFileById(plan_atualizada.getId()))

        //console.log("Planilha cópia excluída")
      }
    }
    resolve("Planilha atualizada")
  })
  
}

function removerLinhas(){
  return new Promise(resolve => {
    while(aba_admp.getRange(aba_admp.getMaxRows(),5).getValue() === null || aba_admp.getRange(aba_admp.getMaxRows(),5).getValue() === ""){
      aba_admp.deleteRow(aba_admp.getMaxRows())
    }
    resolve("Linhas vazias do fim removidas")
  })
}

function padronizar() {
  return new Promise(resolve => {

    const range = "A2:A"
    const range_destino = "b2:b" /* range*/
    const aba_class_gespe = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("class")

    //for(let linha = 2; linha <= 3738; linha++){
      const cargos = aba_class_gespe.getRange(range).getValues()
      
      
      const letras = ["A","E",'I','O','U',"C"]
      
      const erradas = [["Á","Â","Ã"],["É","Ê"],["Í"],["Ó","Ô","Õ"],["Ú"],["Ç"]]
      
      const cargosMai = cargos.map(cargo => {
        var cargoMai = cargo.toString().toUpperCase()
        letras.forEach((substituirPor,i) => {
          erradas[i].forEach(substituirEsta => {
            cargoMai = cargoMai.replace(substituirEsta,substituirPor)
          })
        })

        return [cargoMai]
      })

      aba_class_gespe.getRange(range_destino).setValues(cargosMai)
      
      
    //}

    
    //console.log(coluna_cargo[0])

    resolve("Planilha padronizada")
  })
}

function passarClass(){
  return new Promise(resolve => {
    const lotacoes = aba_admp.getRange("AM2:AM").getValues().map(lotacao=>{
      return [lotacao[0].trim().toUpperCase().replace("C.S.","CENTRO DE SAÚDE")]
    })
    aba_class.getRange("a2:a").clear()
    aba_class.getRange(2,1,lotacoes.length).setValues(lotacoes)
    const ult_linha = aba_class.getRange("a1").getValue()+1
    const classificacoes = aba_class.getRange("b2:c"+ult_linha).getValues()
    aba_admp.getRange("bb1:bc1").setValues([["CLASSIFICAÇÃO","ÁREA DE ATUAÇÃO"]])
    aba_admp.getRange("bb2:bc").setValues(classificacoes)
    resolve("Classificação atualizada")
  })
  
}

function popularEspecialidades(){ 
  return new Promise(resolve => {
    const arr_cargos = aba_admp.getRange("AB2:AB").getValues()
    const arr_esps = aba_admp.getRange("Af2:Af").getValues()

    const arr_esps_corrigidas = arr_esps.map((esp,i) => {
      if(esp[0] === null || esp[0] === ""){
        return arr_cargos[i]
      }
      return arr_esps[i]
    })

    aba_admp.getRange("Af2:Af").setValues(arr_esps_corrigidas)
    
    resolve("Especialidades atualizadas")
  })
  
}

