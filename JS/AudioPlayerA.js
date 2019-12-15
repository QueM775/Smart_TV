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
  console.log('0~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~');
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
    fnRunPlayer(link, objAudioPlayer[0]);
  });
  fnShowCurrentSongFileName();
  /*************************************/
  /* Jump to the next song in the list */
  /*************************************/
  console.log('1~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~');
  objAudioPlayer[0].addEventListener('ended',function(e){
    current++;
    console.log('2~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~');
    if(current > (len)){
      current = 0;
    }
    link = objPlayList.find('a')[current];    
    fnRunPlayer($(link),objAudioPlayer[0]);
  });
}

function fnRunPlayer(link, oPlayerAudio){
  console.log('2~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~');
  oPlayerAudio.src = link.attr('href');
  par = link.parent();
  par.addClass('active').siblings().removeClass('active');
  oPlayerAudio.load();
  oPlayerAudio.play();
  if (document.getElementById('idCurrentSongTitle') !== null){
    fnShowCurrentSongFileName();
  }
}


function fnShowCurrentSongFileName(){
  const sSongNameLong = objAudioPlayer[0].src;
  if (sSongNameLong.length > 0){
    const arrSongName   = sSongNameLong.split("/");
    console.log('arrSongName=' + arrSongName[arrSongName.length - 1]);
    let sSongNameShort = arrSongName[arrSongName.length - 1];
    sSongNameShort = sSongNameShort.replace(/_/g," ");
    sSongNameShort = sSongNameShort.replace(/-/g," ");
    sSongNameShort = sSongNameShort.replace(/.mp3/g,"");
    document.getElementById('idCurrentSongTitle').innerHTML = sSongNameShort;
  }
}