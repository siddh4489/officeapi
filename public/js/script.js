'use strict';

const socket = io();

const outputYou = document.querySelector('.output-you');
const outputBot = document.querySelector('.output-bot');
const outputResult = document.querySelector('.output-result');
outputResult.textContent = 'Initial';

const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
const recognition = new SpeechRecognition();

recognition.lang = 'en-US';
recognition.interimResults = false;
recognition.maxAlternatives = 1;

document.querySelector('.botvoice').addEventListener('click', () => {
  recognition.start();
});

recognition.addEventListener('speechstart', () => {
  console.log('Speech has been detected.');
});

recognition.addEventListener('result', (e) => {
  console.log('Result has been detected.');

  let last = e.results.length - 1;
  let text = e.results[last][0].transcript;

  outputYou.textContent = text;
  console.log('Confidence: ' + e.results[0][0].confidence);
  /*if(text == 'rooms'){
     document.location.href = '/'+text;
     outputBot.textContent = 'Done';
  }else if(text == 'calendar'){
    document.location.href = '/'+text;
    outputBot.textContent = 'Done';
  }else if(text == 'mail'){
    document.location.href = '/'+text;
    outputBot.textContent = 'Done';
  }else if(text =='contacts' || text =='contact'){
    
    //document.location.href = '/'+text;
    $.ajax({
	    type: 'GET',
            contentType: 'application/json',
                    url: '/contacts',						
                    success: function(data) {
			jQuery("#result").html(data);    
                    },
	    	   error  : function(err) { 
			   alert('error');
			   alert('error'+err);
                    }

   });
  // socket.emit('bot reply', 'Contacts fetched.');
   //socket.emit('chat message', 'Contacts fetched');
  synthVoice('Contacts fetched.');
  outputBot.textContent = 'Contacts fetched.'	  
  outputResult.textContent = 'Contacts fetched';

  }else */
  //if(text.includes("meeting") || text.includes("event")){
     //document.location.href = '/event?person='+text;
	
     if(text == 'logout'){
	 document.location.href ='/authorize/signout';
     }	
     var point = jQuery(".output-result").text();
     $.ajax({
	    type: 'GET',
            contentType: 'application/json',
                    url: '/event?person='+text+'&stage='+point,					
                    success: function(data) {
			synthVoice(data.bob);
 		        outputBot.textContent = data.bob;
			outputResult.textContent = data.state;    
			jQuery("#result").html(data.consoleoutput);    
                    },
	    	   error  : function(err) { 
			   alert('error');
			   alert('error'+err);
                    }

   });	  
     //outputBot.textContent = 'Done';
  //}  
  console.log('Text--->'+text);
  //socket.emit('chat message', text);
   //outputResult.textContent = 'Siddhraj Here';

});

recognition.addEventListener('speechend', () => {
  recognition.stop();
});

recognition.addEventListener('error', (e) => {
  outputBot.textContent = 'Error: ' + e.error;
});

function synthVoice(text) {
  const synth = window.speechSynthesis;
  const utterance = new SpeechSynthesisUtterance();
  utterance.text = text;
  synth.speak(utterance);
}

socket.on('bot reply', function(replyText) {
  synthVoice(replyText);
  console.log('------ bot-----'+replyText);
  if(replyText == '') replyText = '(No answer...)';
  outputBot.textContent = replyText;
});
