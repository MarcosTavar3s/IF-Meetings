const form= FormApp.openById('1uanVjY5LHhkS1p8y_UJBPySQvcmew7d9i5SSX1_k9BE'); //forms access
const spreadsheet= SpreadsheetApp.openById('1WkzZtzOY9RWQFFX4g4WC4bhH1WXTV_lEB_UUwcO5oPE'); //spreadsheet access
const workss = spreadsheet.getSheetByName('Respostas ao formulário 1');// working spreadsheet

var lastRow=workss.getLastRow();
var sindex; // first index after the last email

const linkBeginnerIFRN = 'https://classroom.google.com/c/NTUyNTUwNjMxNDA4?cjc=m2y7igu'; //link to access Google Classroom - Beginner IFRN
const linkBeginnerCommunity = 'https://classroom.google.com/c/NjQ1NzUxOTcyMDkx?cjc=px3l67b'; //link to access Google Classroom - Beginner External Community

const linkElementaryIFRN = 'https://classroom.google.com/c/NTA1NTAxNTY0MjQ1?cjc=k7pq2le'; //link to access Google Classroom - Elementary IFRN
const linkElementaryCommunity = 'https://classroom.google.com/c/NjQ1NzYxNzI1MTY5?cjc=3gsbf5h'; //link to access Google Classroom - Elementary External Community

const linkConversation = 'https://sites.google.com/view/ifmeetings/%C3%A1rea-do-aluno/conversation';

function myFunction() {

  for(var i = lastRow; i>0; i--){
    if(workss.getRange(i,3).getBackground()!='#ffffff'){
      sindex=i+1;
      break;
    }
  }

  for(sindex; sindex<=lastRow;sindex++){
    if(workss.getRange(sindex,3).getValue()=='Conversação (níveis intermediário ou avançado)') conversation();
    else if(workss.getRange(sindex,3).getValue()=='Iniciante (quero iniciar os meus estudos na língua inglesa)') beginner();
    else if(workss.getRange(sindex,3).getValue()=='Elementar (terminei o curso iniciante do IF meetings ou tenho condições de começar um curso 100% em inglês)')elementary();

    workss.getRange(sindex,3).setBackground('#b6d7a8');
  }


}

function conversation(){
      var email = workss.getRange(sindex,6).getValue();
      var subject=" Welcome to IF Meetings - Conversation Course ";
      var nameEmail= "IF Meetings - Conversation English Course";
      var body = 'Hey! Welcome to IF Meetings\' conversation course!<br><br>'+'&#128205; Our activities and classes will primarily occur through Google Meetings.<br><br>'+'&#127760; <strong>Our class time is available on our website (click the following link):</strong>'+ linkConversation + '<br><br>' + '&#x23F0; Reminder: You are the one who makes your schedule, then <strong> it is not mandatory to attend every class to receive our certificate.</strong><br><br>' + '&#128226; Frequent Q&A: <br> <ul> <li>Q:\'How many hours do I have to attend to get the certificate?\'<ul>A: You have to attend, at least, 15 classes (15 hours).</ul></li><li>Q:\'If I miss a class, would it have a negative impact on getting my certificate? How could I catch up to my classmates?\'<ul>A: Missing a class does not significantly affect the attainment of the certificate. Additionally, conversation classes do not follow a planned schedule, so you would not be at a disadvantage compared to your classmates.</ul></li></ul><br><br>'+'For further explanation, feel free to contact us through email <strong>ifrnmeetings@gmail.com</strong> or instagram <strong>ifrn_meetings</strong><br><br>' + 'Hope to see you soon!<br>Best regards,<br>IF Meetings\' Team.';

      GmailApp.sendEmail(email, subject, "", {htmlBody: body, name:nameEmail, replyTo: "ifrnmeetings@gmail.com"});

  // console.log('conversation'); // debug
}

function elementary(){
   var email = workss.getRange(sindex,6).getValue();
      var subject=" Welcome to IF Meetings - Elementary Course ";
      var nameEmail= "IF Meetings - Elementary English Course";
      var body;

      if(email.includes('ifrn.edu.br')){
       body = 'Olá! Sejam bem-vind@ ao curso Elementary do IF Meetings!<br><br>'+'&#128221; Nossas atividades e toda nossa interação ocorrerá através do google classroom.<br><br>'+ '&#127760; <strong>Basta acessar o seguinte link:</strong> ' + linkElementaryIFRN +'<br><br>&#x23F0; Vale ressaltar que o curso ocorre em módulos assíncronos, em outras palavras, você irá fazer seu próprio horário. Recebendo certificado ao final do semestre respectivos aos módulos concluídos.'+'<br><br>Caso haja qualquer dúvida, entre em contato conosco pelo nosso email <strong>ifrnmeetings@gmail.com</strong> ou instagram <strong>@ifrn_meetings</strong>.' +'<br><br>Esperamos ver você lá!<br>Atenciosamente,<br>Equipe de monitoria do IF Meetings.';
      }
      else{
        body = 'Olá! Sejam bem-vind@ ao curso Elementary do IF Meetings!<br><br>'+'&#128221; Nossas atividades e toda nossa interação ocorrerá através do google classroom.<br><br>'+ '&#127760; <strong>Basta acessar o seguinte link:</strong> ' + linkElementaryCommunity +'<br><br>&#x23F0; Vale ressaltar que o curso ocorre em módulos assíncronos, em outras palavras, você irá fazer seu próprio horário. Recebendo certificado ao final do semestre respectivos aos módulos concluídos.'+'<br><br>Caso haja qualquer dúvida, entre em contato conosco pelo nosso email <strong>ifrnmeetings@gmail.com</strong> ou instagram <strong>@ifrn_meetings</strong>.' +'<br><br>Esperamos ver você lá!<br>Atenciosamente,<br>Equipe de monitoria do IF Meetings.';

      } 

    GmailApp.sendEmail(email, subject, "", {htmlBody: body, name:nameEmail, replyTo: "ifrnmeetings@gmail.com"});
    //  console.log('elementary'); //debug
}

function beginner(){
   var email = workss.getRange(sindex,6).getValue();
      var subject=" Welcome to IF Meetings - Beginner Course ";
      var nameEmail= "IF Meetings - Beginner English Course";
      var body;

      if(email.includes('ifrn.edu.br')){
        body = 'Olá! Sejam bem-vind@ ao curso Beginner do IF Meetings!<br><br>'+'&#128221; Nossas atividades e toda nossa interação ocorrerá através do google classroom.<br><br>'+ '&#127760; <strong>Basta acessar o seguinte link:</strong> ' + linkBeginnerIFRN +'<br><br>&#x23F0; Vale ressaltar que o curso ocorre em módulos assíncronos, em outras palavras, você irá fazer seu próprio horário. Recebendo certificado ao final do semestre respectivos aos módulos concluídos.'+'<br><br>Caso haja qualquer dúvida, entre em contato conosco pelo nosso email <strong>ifrnmeetings@gmail.com</strong> ou instagram <strong>@ifrn_meetings</strong>.' +'<br><br>Esperamos ver você lá!<br>Atenciosamente,<br>Equipe de monitoria do IF Meetings.';
      }
      else{
        body = 'Olá! Sejam bem-vind@ ao curso Beginner do IF Meetings!<br><br>'+'&#128221; Nossas atividades e toda nossa interação ocorrerá através do google classroom.<br><br>'+ '&#127760; <strong>Basta acessar o seguinte link:</strong> ' + linkBeginnerCommunity +'<br><br>&#x23F0; Vale ressaltar que o curso ocorre em módulos assíncronos, em outras palavras, você irá fazer seu próprio horário. Recebendo certificado ao final do semestre respectivos aos módulos concluídos.'+'<br><br>Caso haja qualquer dúvida, entre em contato conosco pelo nosso email <strong>ifrnmeetings@gmail.com</strong> ou instagram <strong>@ifrn_meetings</strong>.' +'<br><br>Esperamos ver você lá!<br>Atenciosamente,<br>Equipe de monitoria do IF Meetings.';
      
      } 

     GmailApp.sendEmail(email, subject, "", {htmlBody: body, name:nameEmail, replyTo: "ifrnmeetings@gmail.com"});
    //  console.log('beginner'); debug
}
