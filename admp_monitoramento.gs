const esta_plan = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]

function atualizar() {
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
}
