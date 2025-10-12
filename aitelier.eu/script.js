// Prosty przełącznik języka PL <-> EN
document.addEventListener('DOMContentLoaded', () => {
  const html = document.documentElement;
  const btn = document.getElementById('langToggle');

  function setLang(lang){
    html.setAttribute('data-lang', lang);
    btn.textContent = lang === 'pl' ? 'EN' : 'PL';
  }

  btn.addEventListener('click', () => {
    const cur = html.getAttribute('data-lang') || 'pl';
    setLang(cur === 'pl' ? 'en' : 'pl');
  });

  setLang('pl'); // domyślnie polski
});
