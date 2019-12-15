/*
This file is for fnAudioPlayer.asp audio player
It switchs to the next MP3 files when current file (in the folder) ends
It jumps back to the first file after the very last file in current folder

  Dependencies:
  This file must be modified in conjunction with
    defualt.asp (CSS and JS)
    fnAudioPlayer.asp
    fnAudioPlayer_template.html
    AudioPlayerA.js
*/
var objAudioPlayer;
var objPlayList;
var tracks;
var current;
init();
function init(){
  current = 0;
  objAudioPlayer = $('#idAudioPlayer');
  objPlayList = $('#playlist');
  tracks = objPlayList.find('li a');
  len = tracks.length - 1;
  objAudioPlayer[0].volume = .50;
  objPlayList.find('a').click(function(e){
    e.preventDefault();
    link = $(this);
    current = link.parent().index();
    run(link, objAudioPlayer[0]);
  });
  objAudioPlayer[0].addEventListener('ended',function(e){
    current++;
    if(current > (len)){
      current = 0;
    }
    link = objPlayList.find('a')[current];    
    run($(link),objAudioPlayer[0]);
  });
}
function run(link, oPlayerAudio){
  oPlayerAudio.src = link.attr('href');
  par = link.parent();
  par.addClass('active').siblings().removeClass('active');
  oPlayerAudio.load();
  oPlayerAudio.play();
}
