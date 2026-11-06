(() => {
'use strict';
const id='G-DPF4FRDSZB', key='bll-analytics-choice-v1';
let started=false;
function saved(){try{return localStorage.getItem(key);}catch(_){return null;}}
function safeUrl(value){try{const u=new URL(value);return u.origin+u.pathname;}catch(_){return '';}}
function start(){
 window['ga-disable-'+id]=false;window.bllAnalyticsAllowed=true;
 if(started){window.gtag('consent','update',{analytics_storage:'granted'});return;}
 started=true;window.dataLayer=window.dataLayer||[];
 window.gtag=function(){window.dataLayer.push(arguments);};
 window.gtag('consent','default',{analytics_storage:'denied',ad_storage:'denied',ad_user_data:'denied',ad_personalization:'denied'});
 window.gtag('consent','update',{analytics_storage:'granted'});
 window.gtag('js',new Date());
 window.gtag('config',id,{allow_google_signals:false,allow_ad_personalization_signals:false,page_location:location.origin+location.pathname,page_referrer:safeUrl(document.referrer),cookie_expires:31536000});
 const script=document.createElement('script');script.async=true;script.src='https://www.googletagmanager.com/gtag/js?id='+id;document.head.appendChild(script);
}
function stop(){
 window.bllAnalyticsAllowed=false;window['ga-disable-'+id]=true;
 if(typeof window.gtag==='function')window.gtag('consent','update',{analytics_storage:'denied'});
 const names=document.cookie.split(';').map(c=>c.trim().split('=')[0]).filter(n=>/^_ga($|_)/.test(n));
 for(const name of names)for(const domain of [null,location.hostname,'.'+location.hostname,'.bostonluxlimos.com'])document.cookie=name+'=; Max-Age=0; Path=/'+(domain?'; Domain='+domain:'')+'; SameSite=Lax';
}
const panel=document.createElement('aside');panel.id='analytics-choice';panel.setAttribute('aria-label','Optional analytics cookies');
panel.innerHTML='<div><strong>A little insight. Your choice.</strong><p>With your permission, Google Analytics helps us understand page visits, call-button clicks and quote interest. It does not record your calls or collect what you type into our forms. <a href="/privacy/#website-analytics">Learn more</a></p></div><div class="analytics-actions"><button type="button" data-choice="no">Decline analytics</button><button type="button" data-choice="yes">Allow analytics</button></div>';
panel.addEventListener('click',event=>{
 const button=event.target.closest('button[data-choice]');if(!button)return;
 const allowed=button.dataset.choice==='yes';try{localStorage.setItem(key,allowed?'granted':'denied');}catch(_){}
 if(allowed)start();else stop();panel.hidden=true;
});document.body.appendChild(panel);
const footer=document.querySelector('.footer-bottom');
if(footer){const button=document.createElement('button');button.type='button';button.className='analytics-preferences';button.textContent='Cookie preferences';button.addEventListener('click',()=>{panel.hidden=false;panel.querySelector('button').focus();});footer.appendChild(button);}
const choice=saved();panel.hidden=choice==='granted'||choice==='denied';if(choice==='granted')start();else stop();
})();