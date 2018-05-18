'use strict';

const socket = io();

const outputYou = document.querySelector('.output-you');
const outputBot = document.querySelector('.output-bot');
const outputResult = document.querySelector('.output-result');
outputResult.textContent = 'Result are Here';

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
  if(text == 'rooms'){
     document.location.href = '/'+text;
     outputBot.textContent = 'Done';
  }else if(text == 'calendar'){
    document.location.href = '/'+text;
    outputBot.textContent = 'Done';
  }else if(text == 'mail'){
    document.location.href = '/'+text;
    outputBot.textContent = 'Done';
  }else if(text =='contacts'){
    document.location.href = '/'+text+'?q=siddhraj';
  }else if(text.includes("meeting") || text.includes("event")){
     document.location.href = '/event?person='+text;
     outputBot.textContent = 'Done';
  }  
  console.log('Text--->'+text);
  socket.emit('chat message', text);
  outputResult.textContent = 'Siddhraj Here';
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
  //outputBot.textContent = replyText;
});
