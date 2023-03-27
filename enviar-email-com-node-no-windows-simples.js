/*

não precisa instalar ninhum modulo para funcionar porque o FS e o CHILD_PROCESS em teoris vem como padrão

*/

const fs = require('fs');
const {
  exec
} = require('child_process');

function cmd(comandoDeCmd = "tree") {
  exec(comandoDeCmd, (err, stdout, stderr) => {
    if (err) {
      console.error(err);
      return;
    }
    console.log(stdout);
  });
}




function enviarEmail(destinatario,assunto,corpoDoEmail="",pathDeArquivo="") {
  let comandos= `
  Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)

objMail.To = "${destinatario}"
objMail.Subject = "${assunto}"
objMail.Body = "${corpoDoEmail}"`;

if(pathDeArquivo != ""){
    comandos = comandos + `\n objMail.Attachments.Add "${pathDeArquivo}"`;
}
comandos = comandos + `

objMail.Send

Set objMail = Nothing
Set objOutlook = Nothing

  `;
  //guardando a copia do historico
  fs.writeFile('sendEmail.txt', comandos, err => {
    if (err) throw err;
  });

  //versao que vai ser executada
  fs.writeFile('tmp.vbs', comandos, err => {
    if (err) throw err;
  });

  //`echo CreateObject("WScript.Shell").SendKeys "%{UP}" > tmp.vbs && cscript tmp.vbs && del tmp.vbs`
  let comandoParaEnviaroEmail = `cscript tmp.vbs && del tmp.vbs`;
  cmd(comandoParaEnviaroEmail)
}

/*
importante na path usar duas barras para o codigo funcionar
importante ter o outlook instalado e logado na maquina
*/
enviarEmail("claudio.santos86@yahoo.com.br","teste de envio de email 27/0/2023 as 11:44","Email enviado com sucesso com arquivo anexo","path\\to\\file")
