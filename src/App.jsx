import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getAuth, onAuthStateChanged, signOut,
  createUserWithEmailAndPassword, signInWithEmailAndPassword,
  sendPasswordResetEmail, signInWithPopup, deleteUser,
  GoogleAuthProvider, OAuthProvider } from "firebase/auth";
import { getFirestore, doc, getDoc, setDoc, deleteDoc } from "firebase/firestore";

/* Firebase config */
const firebaseApp = initializeApp({
  apiKey: "AIzaSyCZTrJTGH8Jh5WBMhMrV39mjKddRj7p78w",
  authDomain: "zafi-524b8.firebaseapp.com",
  projectId: "zafi-524b8",
  storageBucket: "zafi-524b8.firebasestorage.app",
  messagingSenderId: "308516673564",
  appId: "1:308516673564:web:9410954d5fc50fd56667d9"
});
const auth = getAuth(firebaseApp);
const db = getFirestore(firebaseApp);

/* =========================================================================
   ZAFI — finanzas personales con IA
   - Onboarding conversacional (chatbot tipo Claude) que arma categorías + cuentas
   - Captura de ingresos / egresos
   - Auto-categorización: palabras clave locales -> Claude API -> te pregunta
   - Aprende de tus correcciones (guarda palabras clave)
   - Persiste con window.storage (sobrevive entre sesiones)
   ========================================================================= */

/* ----------------------------- estilos ---------------------------------- */
const STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Montserrat:wght@300;400;500;600;700&display=swap');

html{background-color:var(--bg);}
html,body{min-height:100vh;}
body{
  background-color:var(--bg)!important;
}
/* video background - see <video> in JSX */
.cc-video-bg{position:fixed;inset:0;z-index:-1;overflow:hidden;}
.cc-video-bg::after{content:"";position:absolute;inset:0;background:transparent;pointer-events:none;}
.cc-dark .cc-video-bg::after{background:rgba(0,0,0,.35);}
.cc-video-bg video{width:100%;height:180%;object-fit:cover;
  filter:blur(10px);transform:scale(1.06) translateY(var(--parallax-y, 0px));
  will-change:transform;}
#root{background:transparent!important;min-height:100vh;}
.cc-root *{box-sizing:border-box;margin:0;padding:0;}
:root{
  /* Paleta Zafi — glassmorphism plata/azul frío */
  --bg:#DCE1E8;
  --bg-2:#D2D8E1;
  --paper:rgba(255,255,255,.12);
  --paper-solid:#FFFFFF;
  --surface:rgba(255,255,255,.14);
  --surface-2:rgba(255,255,255,.1);
  --surface-3:rgba(20,30,45,.05);
  --ink:#1B2230;
  --bar-fill:#7E8AA0;
  --ink-soft:#6B7585;
  --ink-faint:#8B95A6;
  --line:rgba(30,40,60,.09);
  --line-soft:rgba(30,40,60,.045);
  --green:#1A7A6E;
  --green-2:#22917F;
  --green-soft:rgba(26,122,110,.10);
  --green-glow:rgba(26,122,110,.20);
  --coral:#B5453A;
  --coral-2:#CC5548;
  --coral-soft:rgba(181,69,58,.08);
  --coral-glow:rgba(181,69,58,.12);
  --gold:#8C7FAE;
  --gold-soft:rgba(140,127,174,.10);
  --gold-glow:rgba(140,127,174,.18);

  /* Acentos del orb IA — iridiscente */
  --orb-purple:#A78BFA;
  --orb-blue:#60A5FA;
  --orb-mint:#5EEAD4;

  /* Degradado de acento (barras de categoría, cuenta seleccionada) */
  --accent-grad:linear-gradient(90deg, #0F2A4A 0%, #1D4ED8 45%, #38BDF8 100%);
  --accent-grad-soft:linear-gradient(135deg, rgba(15,42,74,.10) 0%, rgba(56,189,248,.14) 100%);
  --accent-solid:#1D4ED8;

  /* Sombras frías difusas */
  --shadow-xs:0 1px 3px rgba(30,40,60,.04);
  --shadow-sm:0 2px 8px rgba(30,40,60,.06);
  --shadow-md:0 4px 16px rgba(30,40,60,.08);
  --shadow-lg:0 8px 32px rgba(30,40,60,.10);
  --shadow-xl:0 16px 48px rgba(30,40,60,.13);
  --shadow-inset:inset 0 1px 0 rgba(255,255,255,.7);
  --glass:rgba(255,255,255,.12);
  --glass-border:rgba(255,255,255,.55);
  --blur:blur(5px);
}

/* Dark theme override */
.cc-dark{
  --bg:#0D0F14;
  --bg-2:#151820;
  --paper:rgba(255,255,255,.06);
  --paper-solid:#1A1D24;
  --surface:rgba(255,255,255,.08);
  --surface-2:rgba(255,255,255,.05);
  --surface-3:rgba(255,255,255,.03);
  --ink:#E2E6ED;
  --bar-fill:#8892A4;
  --ink-soft:#8892A4;
  --ink-faint:#5A6270;
  --line:rgba(255,255,255,.1);
  --line-soft:rgba(255,255,255,.06);
  --green:#2DD4A8;
  --green-2:#34E0B4;
  --green-soft:rgba(45,212,168,.12);
  --green-glow:rgba(45,212,168,.22);
  --coral:#F87171;
  --coral-2:#FCA5A5;
  --coral-soft:rgba(248,113,113,.1);
  --coral-glow:rgba(248,113,113,.15);
  --gold:#B4A0D6;
  --gold-soft:rgba(180,160,214,.1);
  --gold-glow:rgba(180,160,214,.18);
  --orb-purple:#B49BFF;
  --orb-blue:#78B8FF;
  --orb-mint:#6EF5DE;
  --accent-grad:linear-gradient(90deg, #1A3A6A 0%, #3B7BF7 45%, #60CFFF 100%);
  --accent-grad-soft:linear-gradient(135deg, rgba(30,60,110,.15) 0%, rgba(96,207,255,.12) 100%);
  --accent-solid:#3B7BF7;
  --shadow-xs:0 1px 3px rgba(0,0,0,.2);
  --shadow-sm:0 2px 8px rgba(0,0,0,.25);
  --shadow-md:0 4px 16px rgba(0,0,0,.3);
  --shadow-lg:0 8px 32px rgba(0,0,0,.35);
  --shadow-xl:0 16px 48px rgba(0,0,0,.4);
  --shadow-inset:inset 0 1px 0 rgba(255,255,255,.1);
  --glass:rgba(255,255,255,.07);
  --glass-border:rgba(255,255,255,.15);
}
.cc-dark .cc-sheet{background:rgba(15,18,25,.45);backdrop-filter:blur(28px);-webkit-backdrop-filter:blur(28px);}
.cc-dark .cc-overlay{background:rgba(0,0,0,.5);}
.cc-dark .cc-input{background:rgba(255,255,255,.06);color:var(--ink);border-color:rgba(255,255,255,.1);}
.cc-dark .cc-input:focus{background:rgba(255,255,255,.1);border-color:rgba(255,255,255,.2);}
.cc-dark .cc-input::placeholder{color:rgba(255,255,255,.3);}
.cc-dark .cc-btn{background:rgba(255,255,255,.08);color:var(--ink);border-color:rgba(255,255,255,.1);}
.cc-dark .cc-btn-primary,.cc-dark .cc-btn-green{background:var(--green);color:#fff;border-color:var(--green);}
.cc-dark .cc-tab{background:rgba(255,255,255,.05);color:var(--ink-soft);}
.cc-dark .cc-tab.on{background:rgba(255,255,255,.12);color:var(--ink);}
.cc-dark .cc-acc-card{background:rgba(255,255,255,.06);border-color:rgba(255,255,255,.1);}
.cc-dark .cc-acc-card.on{background:rgba(255,255,255,.14);border-color:var(--glass-border);}
.cc-dark .cc-toast{background:linear-gradient(160deg,#1a1d24,#2a2d36);color:#E2E6ED;}
.cc-dark .cc-bottomnav-inner{background:rgba(15,17,22,.7);}
.cc-dark .cc-card{background:rgba(255,255,255,.05);}
.cc-dark .cc-bubble.bot{background:rgba(255,255,255,.08);border-color:rgba(255,255,255,.12);}
.cc-dark .cc-bubble.me{background:rgba(255,255,255,.12);border-color:rgba(255,255,255,.15);}
.cc-dark .cc-grip{background:rgba(255,255,255,.2);}
.cc-root{
  font-family:'Montserrat',-apple-system,sans-serif;
  font-weight:300;
  font-size:14px;
  color:var(--ink); background:transparent;
  min-height:100vh; width:100%;
  position:relative;
  isolation:isolate;
  -webkit-font-smoothing:antialiased;
  position:relative;
}
.cc-bg-wave{display:none;}


.cc-num{font-variant-numeric:tabular-nums;font-feature-settings:"tnum";}
.cc-emoji{}
.cc-serif{font-family:'Fraunces',serif;letter-spacing:-.018em;font-feature-settings:"ss01";}

.cc-wrap{max-width:760px;margin:0 auto;padding:4px 16px 130px;}

/* ============== TOPBAR ============== */
.cc-top{position:sticky;top:0;z-index:30;
  background:linear-gradient(to bottom, var(--bg) 0%, rgba(220,225,235,.7) 50%, transparent 100%);
  padding:calc(14px + env(safe-area-inset-top)) 20px 30px;
  transition:.2s ease;}
.cc-top.scrolled{padding-top:calc(9px + env(safe-area-inset-top));padding-bottom:24px;
  background:linear-gradient(to bottom, var(--bg) 0%, rgba(220,225,235,.6) 40%, transparent 100%);
  border-bottom:none;}
.cc-dark .cc-top{background:linear-gradient(to bottom, rgba(13,15,20,.9) 0%, rgba(13,15,20,.5) 50%, transparent 100%);}
.cc-dark .cc-top.scrolled{background:linear-gradient(to bottom, rgba(13,15,20,.85) 0%, rgba(13,15,20,.4) 40%, transparent 100%);}
.cc-top-inner{max-width:760px;margin:0 auto;}

/* logo centrado estilo Canva */
.cc-zafi-wordmark{font-family:'Fraunces',serif;font-weight:400;font-size:16px;
  letter-spacing:-.03em;text-align:center;color:var(--ink);
  font-feature-settings:"ss01"; margin-bottom:12px; transition:.2s;}
.cc-top.scrolled .cc-zafi-wordmark{margin-bottom:6px;font-size:14px;}

/* fila de perfil: avatar + nombre + plan */
.cc-profile-row{display:flex;align-items:center;gap:10px;margin-bottom:12px;transition:.2s;}
.cc-top.scrolled .cc-profile-row{margin-bottom:6px;}
.cc-avatar{width:38px;height:38px;border-radius:50%;flex-shrink:0;
  background:var(--ink);color:#fff;
  display:flex;align-items:center;justify-content:center;
  font-family:'Fraunces',serif;font-weight:600;font-size:15px;
  box-shadow:var(--shadow-sm);transition:.2s;}
.cc-top.scrolled .cc-avatar{width:30px;height:30px;font-size:13px;}
.cc-profile-name{font-family:'Fraunces',serif;font-weight:600;font-size:18px;
  letter-spacing:-.03em;line-height:1.2;color:var(--ink);transition:.2s;}
.cc-top.scrolled .cc-profile-name{font-size:15px;}
.cc-profile-plan{font-size:11px;color:var(--ink-soft);font-weight:500;letter-spacing:-.005em;}
.cc-top.scrolled .cc-profile-plan{display:none;}

.cc-masthead{display:flex;align-items:baseline;justify-content:space-between;gap:12px;
  margin-bottom:10px;transition:.2s;}
.cc-top.scrolled .cc-masthead{margin-bottom:6px;}
.cc-masthead-title{font-family:'Fraunces',serif;font-weight:600;font-size:26px;letter-spacing:-.045em;line-height:1;
  display:flex;align-items:center;gap:8px;font-feature-settings:"ss01";}
.cc-masthead-title::before{content:"";width:8px;height:8px;border-radius:50%;background:var(--green);}

/* Logo del onboarding */
.cc-logo{font-family:'Fraunces',serif;font-weight:400;font-size:26px;letter-spacing:-.05em;
  display:flex;align-items:center;gap:8px;color:var(--ink);font-feature-settings:"ss01";}
.cc-logo-dot{display:none;}
.cc-masthead-meta{font-size:10.5px;font-weight:500;color:var(--ink-faint);letter-spacing:.02em;
  font-variant-numeric:tabular-nums;}

.cc-balance-row{display:flex;align-items:flex-end;justify-content:space-between;gap:14px;padding:2px 0 14px;
  transition:.2s;}
.cc-top.scrolled .cc-balance-row{padding:0 0 8px;}
.cc-top.scrolled .cc-balance-value{font-size:20px;}
.cc-balance-display{display:flex;flex-direction:column;gap:2px;min-width:0;}
.cc-balance-label{font-size:11px;font-weight:600;color:var(--ink-soft);letter-spacing:-.005em;}
.cc-balance-value{font-family:'Fraunces',serif;font-weight:500;font-size:32px;line-height:1;letter-spacing:-.035em;
  font-variant-numeric:tabular-nums;transition:font-size .2s;}

/* chip de rango */
.cc-range-chip{display:inline-flex;align-items:center;gap:7px;padding:8px 14px;
  background:var(--glass);border:1px solid var(--glass-border);border-radius:14px;
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  font-family:inherit;font-size:12.5px;font-weight:600;color:var(--ink);cursor:pointer;
  transition:all .15s ease;}
.cc-range-chip:hover{background:rgba(255,255,255,.7);}
.cc-range-chip:active{transform:scale(.98);}
.cc-range-chip .cc-range-arrow{color:var(--ink-faint);font-size:9px;}

/* tabs — segmented control glass (legado, ya no se usa en header) */
.cc-tabs{display:flex;gap:2px;background:rgba(0,0,0,.04);
  border:1px solid var(--glass-border);border-radius:16px;padding:3px;
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  margin-top:4px;}
.cc-tab{flex:1;font-family:inherit;font-size:12.5px;font-weight:600;
  border:none;background:transparent;color:var(--ink-faint);
  padding:8px 6px;border-radius:13px;cursor:pointer;
  transition:all .15s ease;position:relative;}
.cc-tab:hover{color:var(--ink);}
.cc-tab.on{background:rgba(255,255,255,.1);color:var(--ink);
  box-shadow:var(--shadow-sm);}

/* ============== TARJETAS ============== */
.cc-card{background:var(--glass);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  border:1px solid var(--glass-border);
  border-radius:20px;padding:0;box-shadow:var(--shadow-sm);
  transition:.2s ease;}
.cc-card-boxed{background:var(--glass);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  border:1px solid var(--glass-border);
  border-radius:20px;padding:16px;box-shadow:var(--shadow-sm);}
.cc-card-section{background:transparent;border:none;border-radius:0;padding-top:8px;padding-bottom:4px;
  box-shadow:none;}
.cc-fade{animation:ccUp .4s cubic-bezier(.16,1,.3,1) both;}
@keyframes ccUp{from{opacity:0;}to{opacity:1;}}

/* ============== BOTONES ============== */
.cc-btn{font-family:inherit;font-size:13.5px;font-weight:600;border-radius:14px;border:1px solid var(--glass-border);
  background:var(--glass);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  color:var(--ink);padding:10px 17px;cursor:pointer;
  transition:all .15s ease;box-shadow:none;}
.cc-btn:hover{background:rgba(255,255,255,.7);}
.cc-btn:active{transform:scale(.98);}
.cc-btn-primary{background:var(--ink);color:#fff;border-color:transparent;backdrop-filter:none;-webkit-backdrop-filter:none;}
.cc-btn-primary:hover{background:#222;}
.cc-btn-green{background:var(--ink);color:#fff;
  border-color:transparent;backdrop-filter:none;-webkit-backdrop-filter:none;}
.cc-btn-green:hover{background:#222;}
.cc-btn-green:active{transform:scale(.98);}
.cc-btn:disabled{opacity:.4;cursor:not-allowed;transform:none;}
.cc-btn:disabled:hover{background:var(--paper);border-color:var(--line);}

/* ============== INPUTS ============== */
.cc-input,.cc-select{font-family:inherit;font-size:14.5px;width:100%;padding:12px 16px;border-radius:16px;
  border:1px solid var(--glass-border);background:var(--glass);color:var(--ink);outline:none;
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  transition:border-color .15s ease;}
.cc-input:focus,.cc-select:focus{border-color:rgba(0,0,0,.15);background:rgba(255,255,255,.7);}
.cc-label{font-size:11.5px;font-weight:600;color:var(--ink-soft);margin-bottom:6px;display:block;letter-spacing:-.005em;}

/* ============== CHIPS — píldoras suaves ============== */
.cc-chip{font-family:inherit;font-size:12.5px;font-weight:600;
  border:1px solid var(--line);background:var(--paper);color:var(--ink);
  padding:7px 13px;border-radius:999px;cursor:pointer;
  transition:all .2s cubic-bezier(.2,.7,.2,1);
  box-shadow:var(--shadow-xs);}
.cc-chip:hover{border-color:var(--gold);transform:translateY(-1px);
  box-shadow:var(--shadow-sm);}

/* ============== BOTTOM NAV ============== */
.cc-bottomnav{position:fixed;left:0;right:0;bottom:0;z-index:50;
  display:flex;justify-content:center;
  padding:0 16px calc(14px + env(safe-area-inset-bottom));
  pointer-events:none;}
.cc-bottomnav-inner{pointer-events:auto;
  width:100%;max-width:420px;height:62px;
  display:flex;align-items:center;justify-content:space-between;
  background:rgba(255,255,255,.08);border:1px solid var(--glass-border);
  border-radius:31px;padding:0 10px;
  backdrop-filter:blur(3px);-webkit-backdrop-filter:blur(3px);
  box-shadow:var(--shadow-lg);
  position:relative;}
.cc-nav-item{flex:1;display:flex;flex-direction:column;align-items:center;gap:3px;
  border:none;background:transparent;cursor:pointer;
  padding:8px 4px 6px;border-radius:18px;
  color:var(--ink-faint);font-family:inherit;
  transition:color .15s ease, background .15s ease;}
.cc-nav-item.on{color:var(--ink);}
.cc-nav-item:hover{color:var(--ink);}
.cc-nav-icon{display:flex;align-items:center;justify-content:center;line-height:0;}
.cc-nav-icon svg{width:21px;height:21px;display:block;}
.cc-nav-label{font-size:10px;font-weight:600;letter-spacing:.01em;}

/* Orb central de IA — protruye arriba y abajo de la barra, centrado */
.cc-orb-slot{flex:0 0 78px;display:flex;align-items:center;justify-content:center;position:relative;height:100%;}
.cc-orb-btn{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);
  width:78px;height:78px;border-radius:50%;border:none;cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  box-shadow:var(--shadow-lg);
  transition:transform .2s cubic-bezier(.34,1.56,.64,1), box-shadow .2s ease, opacity .2s ease;}
.cc-orb-btn:active{transform:translate(-50%,-50%) scale(.93);}
.cc-orb{width:78px;height:78px;border-radius:50%;position:relative;
  background:
    radial-gradient(circle at 35% 75%, rgba(99,102,241,.95) 0%, rgba(99,102,241,0) 45%),
    radial-gradient(circle at 70% 55%, rgba(96,165,250,.9) 0%, rgba(96,165,250,0) 50%),
    radial-gradient(circle at 55% 25%, rgba(94,234,212,.55) 0%, rgba(94,234,212,0) 40%),
    radial-gradient(circle at 25% 30%, rgba(167,139,250,.85) 0%, rgba(167,139,250,0) 55%),
    radial-gradient(circle at 50% 50%, #7C8BF5 0%, #5B6EE8 100%);
  filter:saturate(1.25);
  border:2px solid rgba(255,255,255,.55);
  box-shadow:inset -6px -8px 16px rgba(40,30,90,.35), inset 6px 6px 14px rgba(255,255,255,.35);
  animation:ccOrbBreathe 4s ease-in-out infinite, ccOrbDrift 8s ease-in-out infinite;}
.cc-orb::before{content:"";position:absolute;inset:0;border-radius:50%;
  background:radial-gradient(circle at 30% 28%, rgba(255,255,255,.85) 0%, rgba(255,255,255,0) 28%);
  animation:ccOrbShine 4s ease-in-out infinite;}
.cc-orb::after{content:"";position:absolute;inset:-10px;border-radius:50%;z-index:-1;
  background:radial-gradient(circle, rgba(124,139,245,.55) 0%, rgba(167,139,250,.25) 45%, rgba(124,139,245,0) 70%);
  animation:ccOrbGlow 4s ease-in-out infinite;}
@keyframes ccOrbBreathe{0%,100%{transform:scale(1);}50%{transform:scale(1.07);}}
@keyframes ccOrbDrift{
  0%,100%{background-position:35% 75%, 70% 55%, 55% 25%, 25% 30%, 50% 50%;filter:saturate(1.25) hue-rotate(0deg);}
  33%{background-position:55% 65%, 45% 70%, 70% 35%, 35% 45%, 50% 50%;filter:saturate(1.4) hue-rotate(-12deg);}
  66%{background-position:25% 60%, 75% 45%, 40% 30%, 50% 25%, 50% 50%;filter:saturate(1.35) hue-rotate(14deg);}}
@keyframes ccOrbShine{0%,100%{opacity:.85;transform:scale(1);}50%{opacity:1;transform:scale(1.04) translate(2px,2px);}}
@keyframes ccOrbGlow{0%,100%{opacity:.6;transform:scale(1);}50%{opacity:1;transform:scale(1.12);}}

/* ============== FAB superior (+) ============== */
.cc-fab-top{position:fixed;top:18px;right:18px;z-index:45;
  width:46px;height:46px;border-radius:50%;
  background:var(--glass);color:var(--ink);border:1px solid var(--glass-border);
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  font-size:24px;font-weight:300;line-height:1;cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  box-shadow:var(--shadow-md);
  transition:all .15s ease;}
.cc-fab-top:hover{background:rgba(255,255,255,.75);box-shadow:var(--shadow-lg);}
.cc-fab-top:active{transform:scale(.93);}
.cc-fab-top.open{transform:rotate(45deg);}
.cc-fab-menu{position:fixed;top:70px;right:18px;z-index:45;display:flex;flex-direction:column;gap:8px;
  align-items:flex-end;animation:ccUp .15s cubic-bezier(.16,1,.3,1);}
.cc-fab-mini{font-family:inherit;font-size:13px;font-weight:600;padding:10px 16px;border-radius:10px;
  background:var(--paper);color:var(--ink);border:1px solid var(--line);cursor:pointer;
  box-shadow:var(--shadow-md);
  display:flex;align-items:center;gap:8px;
  transition:all .15s ease;}
.cc-fab-mini:hover{background:var(--surface-2);}

/* ============== TARJETAS DE CUENTAS ============== */
.cc-acc-card{cursor:pointer;
  border:1px solid var(--glass-border);background:var(--glass);
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  border-radius:18px;padding:13px 14px;min-width:148px;text-align:left;
  position:relative;overflow:hidden;
  transition:all .2s ease;}
.cc-acc-card:hover{background:rgba(255,255,255,.7);}
.cc-acc-card.on{border-color:rgba(29,78,216,.3);
  background:var(--accent-grad-soft);
  box-shadow:0 0 0 1px rgba(29,78,216,.2);}
.cc-acc-card.on .cc-acc-label{color:var(--accent-solid);}
.cc-acc-card.on .cc-acc-icon{background:rgba(29,78,216,.12);}
.cc-acc-icon{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;
  border-radius:7px;background:var(--surface-2);font-size:13px;margin-bottom:7px;
  transition:.2s;}
.cc-acc-label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--ink-faint);}
.cc-acc-name{font-weight:600;font-size:14px;margin:2px 0 7px;letter-spacing:-.012em;color:var(--ink);}
.cc-acc-bal{font-family:'Fraunces',serif;font-weight:500;font-size:19px;letter-spacing:-.03em;line-height:1.05;}
.cc-grad-text{background:var(--accent-grad);-webkit-background-clip:text;background-clip:text;
  color:transparent;-webkit-text-fill-color:transparent;}
.cc-acc-sub{font-size:10.5px;color:var(--ink-soft);margin-top:4px;font-variant-numeric:tabular-nums;font-weight:500;}
.cc-scroll-x{display:flex;gap:8px;overflow-x:auto;padding:4px 2px 10px;scrollbar-width:none;}
.cc-scroll-x::-webkit-scrollbar{display:none;}

/* configurar */
.cc-gear{background:var(--glass);border:1px solid var(--glass-border);border-radius:14px;
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  padding:8px 14px;cursor:pointer;font-size:12.5px;font-weight:600;color:var(--ink-soft);
  display:inline-flex;align-items:center;gap:6px;
  transition:.15s ease;}
.cc-gear:hover{background:rgba(255,255,255,.7);color:var(--ink);}
.cc-gear svg, .cc-range-chip svg{width:14px;height:14px;flex-shrink:0;}

/* fila draggable */
.cc-sortable{padding:11px 13px;border:1px solid var(--line);border-radius:12px;
  background:var(--paper);display:flex;align-items:center;gap:10px;cursor:grab;
  user-select:none;transition:.18s;box-shadow:var(--shadow-xs);}
.cc-sortable:hover{border-color:var(--gold);box-shadow:var(--shadow-sm);}
.cc-sortable.disabled{opacity:.5;background:var(--surface-2);}
.cc-sortable .cc-grip-h{color:var(--ink-soft);font-size:18px;line-height:1;}

/* ============== SEPARADOR DE DÍA ============== */
.cc-day-sep{display:flex;align-items:center;justify-content:space-between;gap:12px;
  padding:13px 0 8px;border-bottom:1px solid var(--line-soft);margin-bottom:2px;
  position:sticky;top:0;background:transparent;
  z-index:5;}
.cc-day-sep:first-child{padding-top:6px;}
.cc-day-num{font-family:'Fraunces',serif;font-weight:500;font-size:19px;line-height:1;letter-spacing:-.03em;
  color:var(--ink);min-width:24px;}
.cc-day-name{flex:1;font-size:10.5px;font-weight:600;color:var(--ink-faint);letter-spacing:.04em;
  text-transform:uppercase;}
.cc-day-totals{display:flex;gap:9px;align-items:center;font-size:11px;
  font-variant-numeric:tabular-nums;font-weight:500;opacity:.7;}
.cc-day-totals .pos{color:var(--green);}
.cc-day-totals .neg{color:var(--coral);}

/* ============== MODAL ============== */
.cc-overlay{position:fixed;inset:0;background:rgba(0,0,0,.25);
  backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);
  z-index:10000;display:flex;align-items:flex-end;justify-content:center;
  animation:ccFadeIn .15s ease;}
@keyframes ccFadeIn{from{opacity:0;}to{opacity:1;}}
.cc-sheet{background:rgba(255,255,255,.7);backdrop-filter:blur(5px);-webkit-backdrop-filter:blur(5px);
  border-radius:24px 24px 0 0;width:100%;max-width:760px;
  min-height:50vh;max-height:92vh;overflow-y:auto;padding:10px 20px 28px;
  animation:ccSheet .3s cubic-bezier(.16,1,.3,1);
  border-top:1px solid rgba(255,255,255,.6);
  box-shadow:0 -4px 24px rgba(0,0,0,.08);}
@keyframes ccSheet{from{transform:translateY(100%);}to{transform:none;}}
.cc-grip{width:36px;height:4px;background:rgba(0,0,0,.12);border-radius:99px;margin:8px auto 16px;cursor:grab;}
.cc-sheet{transition:transform .25s ease;}
.cc-sheet-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;}
.cc-sheet-top h2{font-family:'Fraunces',serif;font-size:21px;font-weight:600;color:var(--ink);}
.cc-sheet-close{width:32px;height:32px;border-radius:50%;border:none;background:var(--surface);
  display:flex;align-items:center;justify-content:center;cursor:pointer;color:var(--ink-soft);
  font-size:18px;line-height:1;transition:.15s ease;flex-shrink:0;}
.cc-sheet-close:hover{background:var(--surface-2);color:var(--ink);}

/* monto grande centrado — estilo "Nueva transacción" */
.cc-amount-display{display:flex;align-items:baseline;justify-content:center;gap:6px;
  padding:18px 0;margin-bottom:6px;width:100%;}
.cc-amount-display .cc-amount-currency{font-family:'Fraunces',serif;font-size:28px;font-weight:500;color:var(--ink-soft);flex-shrink:0;}
.cc-amount-display .cc-amount-mxn{font-family:'Montserrat',sans-serif;font-size:16px;font-weight:300;color:var(--ink-faint);flex-shrink:0;align-self:center;}
.cc-amount-display input{font-family:'Fraunces',serif;font-size:42px;font-weight:600;letter-spacing:-.02em;
  text-align:left;color:var(--ink);background:transparent;border:none;outline:none;
  width:auto;max-width:240px;min-width:30px;font-feature-settings:"tnum";
  -moz-appearance:textfield;appearance:textfield;}
.cc-amount-display input::-webkit-outer-spin-button,
.cc-amount-display input::-webkit-inner-spin-button{-webkit-appearance:none;-moz-appearance:none;appearance:none;margin:0;display:none;}
.cc-amount-display input[type=number]{-moz-appearance:textfield;}
.cc-amount-display input::placeholder{color:var(--ink-soft);opacity:.55;}

/* ============== CHAT ============== */
.cc-bubble{padding:12px 15px;border-radius:18px;font-size:14.5px;line-height:1.5;max-width:84%;
  letter-spacing:-.005em;}
.cc-bubble.bot{background:rgba(255,255,255,.55);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  color:var(--ink);border:1px solid var(--glass-border);
  border-bottom-left-radius:4px;}
.cc-bubble.me{background:rgba(255,255,255,.45);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  color:var(--ink);border:1px solid var(--glass-border);
  border-bottom-right-radius:6px;align-self:flex-end;
  box-shadow:var(--shadow-sm);}
.cc-mic.rec{background:var(--coral)!important;color:#fff!important;border-color:var(--coral)!important;
  animation:ccMicPulse 1.1s ease-in-out infinite;}
@keyframes ccMicPulse{0%,100%{box-shadow:0 0 0 0 rgba(181,69,58,.5);}50%{box-shadow:0 0 0 8px rgba(181,69,58,0);}}

/* ===== Splash ===== */
.cc-splash{position:fixed;inset:0;z-index:99999;display:flex;flex-direction:column;
  align-items:center;justify-content:center;gap:8px;
  background:var(--bg);
  animation:ccSplashOut .4s ease 1.6s forwards;}
@keyframes ccSplashOut{to{opacity:0;visibility:hidden;}}
.cc-splash-word{font-family:'Fraunces',serif;font-weight:400;
  font-size:48px;letter-spacing:-.05em;color:var(--ink);
  opacity:0;transform:translateY(8px);
  animation:ccSplashFade .8s ease .2s forwards;}
@keyframes ccSplashFade{to{opacity:1;transform:translateY(0);}}
.cc-splash-tag{font-family:'Montserrat',sans-serif;font-weight:300;
  font-size:12px;letter-spacing:.18em;text-transform:uppercase;color:var(--ink-faint);
  opacity:0;animation:ccSplashTag .6s ease .7s forwards;}
@keyframes ccSplashTag{to{opacity:.7;}}

.cc-toast{position:fixed;left:50%;transform:translateX(-50%);bottom:96px;z-index:60;
  background:linear-gradient(160deg,#2c2820,var(--ink));color:var(--paper);
  padding:12px 20px;border-radius:14px;font-size:13.5px;font-weight:600;letter-spacing:-.005em;
  box-shadow:0 12px 32px rgba(31,27,20,.32),var(--shadow-inset);
  animation:ccUp .3s cubic-bezier(.2,.7,.2,1);}

.cc-dots span{display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--ink-soft);margin:0 2px;
  animation:ccDot 1s infinite;}
.cc-dots span:nth-child(2){animation-delay:.15s;}
.cc-dots span:nth-child(3){animation-delay:.3s;}
@keyframes ccDot{0%,60%,100%{opacity:.25;transform:translateY(0);}30%{opacity:1;transform:translateY(-3px);}}

/* ============== ENCABEZADOS DE SECCIÓN ============== */
.cc-eyebrow{font-size:11px;font-weight:600;color:var(--ink-soft);letter-spacing:.02em;
  display:flex;align-items:center;gap:8px;margin-bottom:5px;}
.cc-section-title{font-family:'Fraunces',serif;font-weight:500;font-size:22px;
  letter-spacing:-.03em;line-height:1.15;margin-bottom:14px;}

/* selection */
::selection{background:var(--gold-glow);color:var(--ink);}
`;

/* --------------------------- datos por defecto --------------------------- */
const DEFAULT_CATS = [
  { name: "Sueldo", emoji: "💼", type: "income" },
  { name: "Negocio", emoji: "🏢", type: "income" },
  { name: "Freelance", emoji: "💻", type: "income" },
  { name: "Otros ingresos", emoji: "💰", type: "income" },
  { name: "Súper / Despensa", emoji: "🛒", type: "expense" },
  { name: "Restaurantes", emoji: "🍔", type: "expense" },
  { name: "Transporte / Gasolina", emoji: "⛽", type: "expense" },
  { name: "Casa / Renta", emoji: "🏠", type: "expense" },
  { name: "Servicios", emoji: "💡", type: "expense" },
  { name: "Salud", emoji: "🏥", type: "expense" },
  { name: "Entretenimiento", emoji: "🎬", type: "expense" },
  { name: "Suscripciones", emoji: "📱", type: "expense" },
  { name: "Otros gastos", emoji: "📦", type: "expense" },
];

const SEED_KW = {
  "sueldo": ["sueldo", "salario", "nomina", "quincena"],
  "negocio": ["venta", "cliente", "factura"],
  "freelance": ["freelance", "proyecto", "honorarios"],
  "super / despensa": ["super", "despensa", "walmart", "soriana", "calimax", "mercado", "verdura", "fruta", "leche"],
  "restaurantes": ["restaurante", "taco", "tacos", "pizza", "comida", "cena", "almuerzo", "cafe", "starbucks", "sushi", "hamburguesa"],
  "transporte / gasolina": ["gasolina", "uber", "didi", "taxi", "pasaje", "caseta", "estacionamiento", "pemex"],
  "casa / renta": ["renta", "hipoteca", "mantenimiento"],
  "servicios": ["luz", "cfe", "agua", "internet", "telmex", "izzi", "totalplay", "telefono"],
  "salud": ["doctor", "medicina", "farmacia", "consulta", "dentista", "hospital"],
  "entretenimiento": ["cine", "concierto", "boletos", "antro", "juego"],
  "suscripciones": ["netflix", "spotify", "disney", "suscripcion", "prime", "youtube"],
};

/* ------------------------------ utilidades ------------------------------- */
const uid = () => Math.random().toString(36).slice(2, 10);
const today = () => new Date().toISOString().slice(0, 10);

/* ===== Sistema de idiomas (i18n) ===== */
let _lang = (typeof navigator !== "undefined" && navigator.language?.startsWith("es")) ? "es" : "en";
const setAppLang = (l) => { _lang = l; };
const STRINGS = {
  // Tabs
  home: { es: "Inicio", en: "Home" },
  history: { es: "Historial", en: "History" },
  categories: { es: "Categorías", en: "Categories" },
  statistics: { es: "Estadísticas", en: "Statistics" },
  // Header
  weeklyPlan: { es: "Plan semanal", en: "Weekly plan" },
  // Home sections
  yourAccounts: { es: "Tus cuentas", en: "Your accounts" },
  addAccount: { es: "Agregar cuenta", en: "Add account" },
  customize: { es: "Personalizar", en: "Customize" },
  incomeThisMonth: { es: "Ingresos · Este mes", en: "Income · This month" },
  expenseThisMonth: { es: "Gastos · Este mes", en: "Expenses · This month" },
  topExpenses: { es: "Gastos más grandes", en: "Biggest expenses" },
  recentMovements: { es: "Últimos movimientos", en: "Recent movements" },
  expByCategory: { es: "Gastos por categoría", en: "Expenses by category" },
  balance: { es: "Saldo", en: "Balance" },
  // History
  all: { es: "Todos", en: "All" },
  income: { es: "Ingresos", en: "Income" },
  expenses: { es: "Gastos", en: "Expenses" },
  select: { es: "Seleccionar", en: "Select" },
  selected: { es: "seleccionado", en: "selected" },
  selectAll: { es: "Todos", en: "All" },
  none: { es: "Ninguno", en: "None" },
  deleteBtn: { es: "Eliminar", en: "Delete" },
  sortBy: { es: "Ordenar", en: "Sort by" },
  mov: { es: "mov.", en: "mov." },
  touchToEdit: { es: "Toca un movimiento para editarlo · ✕ para eliminarlo", en: "Tap a movement to edit it · ✕ to delete it" },
  noMovements: { es: "No hay movimientos en este periodo.", en: "No movements in this period." },
  // AddModal
  newTransaction: { es: "Nueva transacción", en: "New transaction" },
  editTransaction: { es: "Editar transacción", en: "Edit transaction" },
  expense: { es: "Gasto", en: "Expense" },
  incomeType: { es: "Ingreso", en: "Income" },
  concept: { es: "Concepto", en: "Concept" },
  date: { es: "Fecha", en: "Date" },
  category: { es: "Categoría", en: "Category" },
  account: { es: "Cuenta", en: "Account" },
  save: { es: "Guardar", en: "Save" },
  saveChanges: { es: "Guardar cambios", en: "Save changes" },
  // TopFab menu
  addMovement: { es: "Agregar movimiento", en: "Add movement" },
  manualCapture: { es: "Capturar manual", en: "Manual capture" },
  manualCaptureDesc: { es: "Escribe el movimiento tú mismo", en: "Enter the movement yourself" },
  recurringMovement: { es: "Movimiento recurrente", en: "Recurring movement" },
  recurringDesc: { es: "Se repite automáticamente", en: "Repeats automatically" },
  fromScreenshot: { es: "Desde screenshot", en: "From screenshot" },
  fromScreenshotDesc: { es: "Sube una captura y la IA lo lee", en: "Upload a screenshot and AI reads it" },
  fromExcel: { es: "Desde Excel", en: "From Excel" },
  fromExcelDesc: { es: "Importa una hoja de cálculo", en: "Import a spreadsheet" },
  // Settings
  settings: { es: "Ajustes", en: "Settings" },
  personalInfo: { es: "Información personal", en: "Personal information" },
  name: { es: "Nombre", en: "Name" },
  phone: { es: "Teléfono", en: "Phone" },
  email: { es: "Correo electrónico", en: "Email" },
  language: { es: "Idioma", en: "Language" },
  currency: { es: "Moneda", en: "Currency" },
  notifications: { es: "Notificaciones", en: "Notifications" },
  expenseReminders: { es: "Recordatorios de gastos", en: "Expense reminders" },
  comingSoon: { es: "Próximamente", en: "Coming soon" },
  dataPrivacy: { es: "Datos y privacidad", en: "Data & privacy" },
  exportData: { es: "Exportar mis datos", en: "Export my data" },
  resetApp: { es: "Reiniciar app (borrar todo)", en: "Reset app (delete all)" },
  resetConfirm: { es: "¿Estás seguro? Esto borrará todas tus categorías, cuentas y movimientos.", en: "Are you sure? This will delete all your categories, accounts and movements." },
  cancel: { es: "Cancelar", en: "Cancel" },
  yesDeleteAll: { es: "Sí, borrar todo", en: "Yes, delete all" },
  signOut: { es: "Cerrar sesión", en: "Sign out" },
  signingOut: { es: "Cerrando sesión…", en: "Signing out…" },
  infoUpdated: { es: "Información actualizada", en: "Information updated" },
  chooseAvatar: { es: "Elige tu avatar", en: "Choose your avatar" },
  avatarUpdated: { es: "Avatar actualizado", en: "Avatar updated" },
  avatarRemoved: { es: "Avatar eliminado", en: "Avatar removed" },
  choosePhoto: { es: "Elegir foto", en: "Choose photo" },
  takePhoto: { es: "Tomar foto", en: "Take photo" },
  orUploadPhoto: { es: "O sube tu propia foto", en: "Or upload your own photo" },
  // Recurring
  recurringMovements: { es: "Movimientos recurrentes", en: "Recurring movements" },
  daily: { es: "Diario", en: "Daily" },
  weekly: { es: "Semanal", en: "Weekly" },
  biweekly: { es: "Quincenal", en: "Biweekly" },
  monthly: { es: "Mensual", en: "Monthly" },
  yearly: { es: "Anual", en: "Yearly" },
  manage: { es: "Gestionar", en: "Manage" },
  newRecurring: { es: "Nuevo recurrente", en: "New recurring" },
  pause: { es: "Pausar", en: "Pause" },
  activate: { es: "Activar", en: "Activate" },
  edit: { es: "Editar", en: "Edit" },
  // Stats
  summary: { es: "Resumen", en: "Summary" },
  netFlow: { es: "Flujo neto", en: "Net flow" },
  tapForDetail: { es: "Tocar para detalle ▸", en: "Tap for detail ▸" },
  avgDailyExpense: { es: "Gasto promedio diario", en: "Avg. daily expense" },
  movementsInPeriod: { es: "Movimientos en el periodo", en: "Movements in period" },
  incByCategory: { es: "Ingresos por categoría", en: "Income by category" },
  topSpent: { es: "En lo que más gastaste", en: "Top spending" },
  ofYourExpenses: { es: "de tus gastos", en: "of your expenses" },
  customizeStats: { es: "Personalizar", en: "Customize" },
  customizeStatsTitle: { es: "Personalizar estadísticas", en: "Customize statistics" },
  reorderHint: { es: "Reordena con las flechas y muestra u oculta cada sección.", en: "Reorder with arrows and show or hide each section." },
  bars: { es: "Barras", en: "Bars" },
  pie: { es: "Pastel", en: "Pie" },
  donut: { es: "Dona", en: "Donut" },
  slideToSee: { es: "Desliza sobre la gráfica para ver cualquier día", en: "Slide over the chart to see any day" },
  // Categories tab
  eachAccountHasOwn: { es: "Cada cuenta tiene sus propias categorías.", en: "Each account has its own categories." },
  totalsAreFrom: { es: "Los totales son de", en: "Totals are from" },
  noRecurring: { es: "No tienes movimientos recurrentes. Crea uno para que se registre automáticamente.", en: "No recurring movements. Create one to have it registered automatically." },
  activeCount: { es: "activo", en: "active" },
  autoGenerated: { es: "se generan solos en su fecha", en: "auto-generated on their date" },
  addCategoryTo: { es: "Agregar categoría a", en: "Add category to" },
  noCats: { es: "Sin categorías.", en: "No categories." },
  inPeriod: { es: "en el periodo", en: "in period" },
  // Assistant
  assistant: { es: "Asistente", en: "Assistant" },
  close: { es: "Cerrar", en: "Close" },
  send: { es: "Enviar", en: "Send" },
  releaseToSend: { es: "Suelta para enviar", en: "Release to send" },
  listening: { es: "Escuchando…", en: "Listening…" },
  // General
  loading: { es: "Cargando…", en: "Loading…" },
  user: { es: "Usuario", en: "User" },
  from: { es: "Desde", en: "From" },
  to: { es: "Hasta", en: "To" },
  apply: { es: "Aplicar", en: "Apply" },
  male: { es: "Hombre", en: "Male" },
  female: { es: "Mujer", en: "Female" },
  other: { es: "Otro", en: "Other" },
  paused: { es: "pausado", en: "paused" },
  // Auth screens
  welcomeBack: { es: "Bienvenido de vuelta", en: "Welcome back" },
  signInToContinue: { es: "Inicia sesión para continuar", en: "Sign in to continue" },
  createAccount: { es: "Crear cuenta", en: "Create account" },
  signIn: { es: "Iniciar sesión", en: "Sign in" },
  continueWithEmail: { es: "Continuar con tu correo", en: "Continue with your email" },
  continueWithGoogle: { es: "Continuar con Google", en: "Continue with Google" },
  continueWithApple: { es: "Continuar con Apple", en: "Continue with Apple" },
  yourName: { es: "Tu nombre", en: "Your name" },
  howToCallYou: { es: "¿Cómo te llamamos?", en: "How should we call you?" },
  age: { es: "Edad", en: "Age" },
  gender: { es: "Género", en: "Gender" },
  country: { es: "País", en: "Country" },
  password: { es: "Contraseña", en: "Password" },
  confirmPassword: { es: "Confirmar contraseña", en: "Confirm password" },
  continueBtn: { es: "Continuar", en: "Continue" },
  back: { es: "Atrás", en: "Back" },
  forgotPassword: { es: "¿Olvidaste tu contraseña?", en: "Forgot password?" },
  resetPassword: { es: "Recuperar contraseña", en: "Reset password" },
  resetPasswordDesc: { es: "Escribe tu correo y te mandamos un enlace para crear una nueva contraseña.", en: "Enter your email and we'll send you a link to create a new password." },
  sendLink: { es: "Enviar enlace", en: "Send link" },
  linkSent: { es: "¡Listo! Revisa tu correo.", en: "Done! Check your email." },
  aboutYou: { es: "Sobre ti", en: "About you" },
  aboutYouDesc: { es: "Para personalizar tu experiencia.", en: "To personalize your experience." },
  yourCredentials: { es: "Tus credenciales", en: "Your credentials" },
  credentialsDesc: { es: "Para acceder a tu cuenta en cualquier dispositivo.", en: "To access your account on any device." },
  welcomeToZafi: { es: "Bienvenido a Zafi", en: "Welcome to Zafi" },
  tellUsAboutYou: { es: "Cuéntanos un poco de ti para personalizar tu experiencia.", en: "Tell us a bit about yourself so we can personalize your experience." },
  saving: { es: "Guardando…", en: "Saving…" },
  enterName: { es: "Escribe tu nombre.", en: "Enter your name." },
  enterValidAge: { es: "Escribe una edad válida.", en: "Enter a valid age." },
  selectGender: { es: "Selecciona tu género.", en: "Select your gender." },
  fillAllFields: { es: "Llena todos los campos.", en: "Fill in all fields." },
  passwordsDontMatch: { es: "Las contraseñas no coinciden.", en: "Passwords don't match." },
  passwordMinLength: { es: "Mínimo 6 caracteres en la contraseña.", en: "Password must be at least 6 characters." },
  selectCountry: { es: "Selecciona tu país.", en: "Select your country." },
  somethingWentWrong: { es: "Algo salió mal. Intenta de nuevo.", en: "Something went wrong. Try again." },
  countryOfResidence: { es: "País de residencia", en: "Country of Residence" },
  countryDesc: { es: "Necesitamos algunos datos personales para verificar tu elegibilidad.", en: "We need to collect a few personal details to verify you're elegible for an account" },
  getStarted: { es: "Comenzar", en: "Get started" },
  legalConsent: { es: "Al tocar Comenzar, aceptas los", en: "By tapping Get started, you agree to Zafi's" },
  termsAndConditions: { es: "Términos y condiciones", en: "Terms and Conditions" },
  privacyPolicy: { es: "Política de privacidad", en: "Privacy Policy" },
  legalConsentEnd: { es: "y declaras que actúas en tu propio nombre.", en: "and represent that you are acting on your own behalf." },
};
const t = (key) => (STRINGS[key] || {})[_lang] || (STRINGS[key] || {}).es || key;

/* ===== Motor de movimientos recurrentes =====
 * Una regla recurrente vive en config.recurring[] con forma:
 * { id, type:"expense"|"income", amount, description, accountId, categoryId,
 *   freq:"daily"|"weekly"|"biweekly"|"monthly"|"yearly", startDate:"YYYY-MM-DD",
 *   lastRun:"YYYY-MM-DD"|null, active:true }
 */
const FREQ_LABELS_FN = () => ({
  daily: t("daily"), weekly: t("weekly"), biweekly: t("biweekly"),
  monthly: t("monthly"), yearly: t("yearly"),
});
// Una regla está activa salvo que esté pausada explícitamente (compat: asistente usa `paused`, modal usa `active`)
const isRecActive = (r) => r.active !== false && r.paused !== true;

/* ===== Dictado por voz (Web Speech API) ===== */
const SpeechRec = typeof window !== "undefined"
  ? (window.SpeechRecognition || window.webkitSpeechRecognition)
  : null;
const voiceSupported = !!SpeechRec;
const isoToDate = (iso) => { const [y, m, d] = iso.split("-").map(Number); return new Date(y, m - 1, d); };
const dateToIso = (dt) => {
  const y = dt.getFullYear(); const m = String(dt.getMonth() + 1).padStart(2, "0"); const d = String(dt.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
};
// avanza una fecha según la frecuencia
const advanceDate = (dt, freq) => {
  const n = new Date(dt);
  if (freq === "daily") n.setDate(n.getDate() + 1);
  else if (freq === "weekly") n.setDate(n.getDate() + 7);
  else if (freq === "biweekly") n.setDate(n.getDate() + 14);
  else if (freq === "monthly") n.setMonth(n.getMonth() + 1);
  else if (freq === "yearly") n.setFullYear(n.getFullYear() + 1);
  return n;
};
// genera todos los movimientos pendientes de una regla, desde startDate (o lastRun) hasta hoy
const runRecurringRules = (recurring) => {
  if (!recurring || !recurring.length) return { newTxs: [], updatedRecurring: recurring || [] };
  const todayIso = today();
  const todayD = isoToDate(todayIso);
  const newTxs = [];
  const updatedRecurring = recurring.map((r) => {
    if (!isRecActive(r)) return r;
    // primera fecha pendiente
    let cursor;
    if (r.lastRun) {
      cursor = advanceDate(isoToDate(r.lastRun), r.freq);
    } else {
      cursor = isoToDate(r.startDate);
    }
    let lastRun = r.lastRun;
    let guard = 0;
    while (cursor <= todayD && guard < 1000) {
      const iso = dateToIso(cursor);
      newTxs.push({
        id: uid(),
        type: r.type,
        amount: r.amount,
        description: r.description,
        categoryId: r.categoryId || null,
        accountId: r.accountId,
        date: iso,
        recurringId: r.id,
      });
      lastRun = iso;
      cursor = advanceDate(cursor, r.freq);
      guard++;
    }
    return { ...r, lastRun };
  });
  return { newTxs, updatedRecurring };
};

/* ===== Detector de patrones frecuentes (últimos 2 meses) ===== */
function detectFrequentPatterns(txs, config) {
  const now = new Date();

  // Ventana 1: últimos 2 meses
  const twoMonthsAgo = new Date();
  twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);
  const cutoffRecent = dateToIso(twoMonthsAgo);
  const cutoffNow = dateToIso(now);

  // Ventana 2: misma ventana del año pasado (mes actual ±1 mes, año anterior)
  const lastYearStart = new Date(now.getFullYear() - 1, now.getMonth() - 1, 1);
  const lastYearEnd = new Date(now.getFullYear() - 1, now.getMonth() + 2, 0);
  const cutoffLYStart = dateToIso(lastYearStart);
  const cutoffLYEnd = dateToIso(lastYearEnd);

  const recent = (txs || []).filter(t => {
    if (t.recurringId) return false;
    const d = t.date;
    // en ventana reciente O en ventana del año pasado
    return (d >= cutoffRecent && d <= cutoffNow) || (d >= cutoffLYStart && d <= cutoffLYEnd);
  });

  // agrupar por descripción normalizada
  const groups = {};
  recent.forEach(t => {
    const key = (t.description || "").toLowerCase().trim();
    if (!key || key.length < 2) return;
    if (!groups[key]) groups[key] = [];
    groups[key].push(t);
  });

  // encontrar patrones que se repiten 2+ veces con montos similares
  const patterns = [];
  Object.entries(groups).forEach(([key, txArr]) => {
    if (txArr.length < 2) return;
    const avgAmt = txArr.reduce((s, t) => s + t.amount, 0) / txArr.length;
    const last = txArr.sort((a, b) => b.date.localeCompare(a.date))[0];
    const cat = config.categories.find(c => c.id === last.categoryId);
    // detectar si hay ocurrencias en ambos años
    const hasRecent = txArr.some(t => t.date >= cutoffRecent);
    const hasLastYear = txArr.some(t => t.date >= cutoffLYStart && t.date <= cutoffLYEnd);
    patterns.push({
      description: last.description,
      amount: Math.round(avgAmt),
      type: last.type,
      categoryId: last.categoryId,
      categoryName: cat ? cat.name : null,
      categoryEmoji: cat ? cat.emoji : null,
      accountId: last.accountId,
      count: txArr.length,
      lastDate: last.date,
      annual: hasLastYear && hasRecent, // patrón anual (aparece en ambos años)
    });
  });

  // ordenar: anuales primero, luego por frecuencia
  return patterns.sort((a, b) => (b.annual ? 1 : 0) - (a.annual ? 1 : 0) || b.count - a.count).slice(0, 10);
}

const fmtBare = (n) => {
  const v = n || 0;
  const hasDecimals = v % 1 !== 0;
  return new Intl.NumberFormat("es-MX", { style: "currency", currency: "MXN",
    minimumFractionDigits: hasDecimals ? 2 : 0, maximumFractionDigits: hasDecimals ? 2 : 0 }).format(v);
};
const fmt = (n) => `${fmtBare(n)} mxn`;
const fmtMxn = (n) => fmt(n);
const norm = (s) =>
  (s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9 ]/g, " ");
const monthLabel = (mk) => {
  const [y, m] = mk.split("-");
  const names = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  return `${names[+m - 1]} ${y}`;
};
const DAY_NAMES = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
const MONTH_NAMES_FULL = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
const dayLabel = (isoDate) => {
  // YYYY-MM-DD → "jueves 19 de mayo 2026" sin desfase de zona horaria
  const [y, m, d] = isoDate.split("-").map(Number);
  if (!y || !m || !d) return isoDate;
  const date = new Date(y, m - 1, d);
  const todayD = new Date(today());
  const yest = new Date(todayD); yest.setDate(yest.getDate() - 1);
  const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  if (sameDay(date, todayD)) return "Hoy";
  if (sameDay(date, yest)) return "Ayer";
  return `${DAY_NAMES[date.getDay()]} ${d} de ${MONTH_NAMES_FULL[m - 1]} ${y}`;
};

// Para el estilo editorial: número del día grande + descripción a un lado
const dayParts = (isoDate) => {
  const [y, m, d] = isoDate.split("-").map(Number);
  if (!y || !m || !d) return { num: "—", desc: isoDate, isSpecial: false };
  const date = new Date(y, m - 1, d);
  const todayD = new Date(today());
  const yest = new Date(todayD); yest.setDate(yest.getDate() - 1);
  const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  if (sameDay(date, todayD)) return { num: String(d), desc: "HOY", isSpecial: true };
  if (sameDay(date, yest)) return { num: String(d), desc: "AYER", isSpecial: true };
  const mon = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"][m-1];
  const day = ["DOMINGO","LUNES","MARTES","MIÉRCOLES","JUEVES","VIERNES","SÁBADO"][date.getDay()];
  return { num: String(d), desc: `${day} · ${mon} ${y}`, isSpecial: false };
};
const STOP = new Set(["para", "con", "los", "las", "del", "una", "uno", "este", "esta", "que", "por", "mas",
  "muy", "pago", "compra", "gasto", "abono", "the", "and"]);

function extractKW(desc) {
  return norm(desc).split(/\s+/).filter((w) => w.length >= 4 && !STOP.has(w));
}

/* detección de duplicados: ¿existe ya un movimiento equivalente?
   - misma cuenta, mismo tipo, mismo monto (a 2 decimales)
   - fecha dentro de ±3 días (los bancos a veces ponen fecha de operación vs aplicación)
   - descripción "parecida": comparten al menos una palabra clave significativa,
     o las descripciones normalizadas son iguales */
function findDuplicate(candidate, existing) {
  const candDate = new Date(candidate.date).getTime();
  const candKws = new Set(extractKW(candidate.description));
  const candNorm = norm(candidate.description).trim();
  const candAmt = Math.round(candidate.amount * 100);
  for (const t of existing) {
    if (t.accountId !== candidate.accountId) continue;
    if (t.type !== candidate.type) continue;
    if (Math.round(t.amount * 100) !== candAmt) continue;
    const tDate = new Date(t.date).getTime();
    const diffDays = Math.abs((tDate - candDate) / 86400000);
    if (diffDays > 3) continue;
    const tNorm = norm(t.description).trim();
    if (candNorm && tNorm && candNorm === tNorm) return t;
    const tKws = extractKW(t.description);
    if (tKws.some((kw) => candKws.has(kw))) return t;
    // si una descripción está vacía pero monto+fecha+cuenta+tipo coinciden, también es duplicado probable
    if (!candNorm || !tNorm) return t;
  }
  return null;
}

/* saldos: saldo inicial + movimientos */
function accountBalance(config, txs, accId) {
  const acc = config.accounts.find((a) => a.id === accId);
  const init = acc ? acc.initialBalance || 0 : 0;
  return init + txs.filter((t) => t.accountId === accId)
    .reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
}
function grandTotal(config, txs) {
  const init = config.accounts.reduce((s, a) => s + (a.initialBalance || 0), 0);
  return init + txs.reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
}

/* Para estadísticas: separa entre movimientos reales (afectan ingresos/gastos)
   y pass-through (dinero de paso). Los pass-through con el mismo groupId
   forman un grupo: sumamos ingresos del grupo, restamos gastos del grupo,
   la diferencia se cuenta como ingreso/gasto real (la "propina" o "lo que pusiste").
   Soporta cualquier combinación: 1-1, 1-n, n-1, n-m. */
function statTxs(txs) {
  const real = [];
  const synthetic = [];
  const used = new Set();
  // 1. agrupar pass-through por groupId
  const groups = new Map();
  for (const t of txs) {
    if (!t.passThrough) { real.push(t); continue; }
    if (!t.groupId) { used.add(t.id); continue; } // pass-through sin grupo → ignorar en stats
    if (!groups.has(t.groupId)) groups.set(t.groupId, []);
    groups.get(t.groupId).push(t);
  }
  // 2. para cada grupo: sumar ingresos vs gastos
  for (const [gid, members] of groups) {
    let totalIn = 0, totalOut = 0;
    members.forEach((m) => {
      used.add(m.id);
      if (m.type === "income") totalIn += m.amount;
      else totalOut += m.amount;
    });
    const diff = totalIn - totalOut;
    if (Math.abs(diff) < 0.01) continue;
    // fecha del sintético: la más reciente del grupo
    const latest = members.reduce((a, b) => (a.date >= b.date ? a : b));
    // descripción del sintético: tomar la de algún miembro
    const sampleDesc = members.find((m) => m.description)?.description || "reembolso";
    synthetic.push({
      id: "synth-" + gid,
      type: diff >= 0 ? "income" : "expense",
      amount: Math.abs(diff),
      description: `Diferencia neta · ${sampleDesc}`,
      categoryId: null,
      accountId: members[0].accountId,
      date: latest.date,
    });
  }
  return { real, synthetic, all: [...real, ...synthetic] };
}

/* ===================== RANGO TEMPORAL GLOBAL ============================ */
/* El rango global vive en config.dateRange con forma:
     { preset: "today"|"week"|"month"|"last-month"|"3m"|"6m"|"year"|"last-year"|"all"|"custom",
       from?: "YYYY-MM-DD", to?: "YYYY-MM-DD" }
   Si preset !== "custom", from/to se calculan dinámicamente con resolveRange().
*/
const DEFAULT_RANGE = { preset: "month" };

function resolveRange(range) {
  const r = range || DEFAULT_RANGE;
  const t = today();
  const td = new Date(t + "T12:00:00");
  const iso = (d) => d.toISOString().slice(0, 10);
  function firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1, 12); }
  function lastOfMonth(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 0, 12); }
  switch (r.preset) {
    case "today": return { from: t, to: t };
    case "week": {
      const dow = td.getDay();
      const monOff = dow === 0 ? -6 : 1 - dow;
      const mon = new Date(td); mon.setDate(mon.getDate() + monOff);
      const sun = new Date(mon); sun.setDate(sun.getDate() + 6);
      return { from: iso(mon), to: iso(sun) };
    }
    case "month": return { from: iso(firstOfMonth(td)), to: t };
    case "last-month": {
      const prev = new Date(td.getFullYear(), td.getMonth() - 1, 15, 12);
      return { from: iso(firstOfMonth(prev)), to: iso(lastOfMonth(prev)) };
    }
    case "3m": {
      const start = new Date(td.getFullYear(), td.getMonth() - 2, 1, 12);
      return { from: iso(start), to: t };
    }
    case "6m": {
      const start = new Date(td.getFullYear(), td.getMonth() - 5, 1, 12);
      return { from: iso(start), to: t };
    }
    case "year": return { from: `${td.getFullYear()}-01-01`, to: t };
    case "last-year": {
      const y = td.getFullYear() - 1;
      return { from: `${y}-01-01`, to: `${y}-12-31` };
    }
    case "all": return { from: "1970-01-01", to: t };
    case "custom":
      return { from: r.from || t, to: r.to || t };
    default: return { from: iso(firstOfMonth(td)), to: t };
  }
}

function rangeLabel(range) {
  const r = range || DEFAULT_RANGE;
  if (r.preset === "today") return "Hoy";
  if (r.preset === "week") return "Esta semana";
  if (r.preset === "month") return "Este mes";
  if (r.preset === "last-month") return "Mes pasado";
  if (r.preset === "3m") return "Últimos 3 meses";
  if (r.preset === "6m") return "Últimos 6 meses";
  if (r.preset === "year") return "Este año";
  if (r.preset === "last-year") return "Año pasado";
  if (r.preset === "all") return "Todo el historial";
  // custom
  const { from, to } = resolveRange(r);
  if (from === to) return prettyDate(from);
  return `${prettyDate(from)} – ${prettyDate(to)}`;
}

function prettyDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return iso;
  const M = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  return `${d} ${M[m - 1]} ${y}`;
}

/* filtra txs según rango (los movimientos fuera del rango se omiten) */
function txsInRange(txs, range) {
  const { from, to } = resolveRange(range);
  return txs.filter((t) => t.date >= from && t.date <= to);
}

/* elige granularidad sensata según el ancho del rango */
function rangeGranularity(range) {
  const { from, to } = resolveRange(range);
  const days = Math.max(1, Math.round((new Date(to) - new Date(from)) / 86400000) + 1);
  if (days <= 14) return "day";
  if (days <= 92) return "week";
  if (days <= 730) return "month";
  return "month";
}

/* convierte una spec de gráfica del asistente en el payload que consume MiniChart */
function buildChart(spec, config, txs) {
  /* spec esperada del asistente:
     { kind: "line"|"bars"|"pie",
       title?: "string",
       groupBy: "day"|"week"|"month"|"category"|"account",
       metric: "income"|"expense"|"both"|"net"|"expense_by_category",
       dateFrom: "YYYY-MM-DD", dateTo: "YYYY-MM-DD",
       accountId?: "..." | null,
       includePassThrough?: false  // default false }
  */
  if (!spec || !spec.kind) return null;
  const kind = spec.kind;
  const includePT = !!spec.includePassThrough;
  const metric = spec.metric || "both";
  const groupBy = spec.groupBy || "month";
  // rango por defecto: mes actual
  let dateFrom = spec.dateFrom || today().slice(0, 7) + "-01";
  let dateTo = spec.dateTo || today();
  if (dateFrom > dateTo) { const tmp = dateFrom; dateFrom = dateTo; dateTo = tmp; }

  // filtrar
  let pool = txs.filter((t) => t.date >= dateFrom && t.date <= dateTo);
  if (spec.accountId) pool = pool.filter((t) => t.accountId === spec.accountId);
  // aplicar statTxs (excluye pass-through, agrega diferencias netas)
  pool = includePT ? pool : statTxs(pool).all;

  // helpers de buckets
  const buckets = [];
  if (groupBy === "day") {
    const d0 = new Date(dateFrom + "T12:00:00");
    const d1 = new Date(dateTo + "T12:00:00");
    for (let d = new Date(d0); d <= d1; d.setDate(d.getDate() + 1)) {
      const k = d.toISOString().slice(0, 10);
      buckets.push({ key: k, label: `${d.getDate()}` });
    }
  } else if (groupBy === "week") {
    // semanas lunes-domingo
    const d0 = new Date(dateFrom + "T12:00:00");
    const d1 = new Date(dateTo + "T12:00:00");
    const startMon = new Date(d0);
    const dow = startMon.getDay(); // 0 dom, 1 lun
    const offset = (dow === 0 ? -6 : 1 - dow);
    startMon.setDate(startMon.getDate() + offset);
    let n = 1;
    for (let s = new Date(startMon); s <= d1; s.setDate(s.getDate() + 7)) {
      const end = new Date(s); end.setDate(end.getDate() + 6);
      buckets.push({
        key: s.toISOString().slice(0, 10) + "..." + end.toISOString().slice(0, 10),
        label: `S${n}`,
        startISO: s.toISOString().slice(0, 10),
        endISO: end.toISOString().slice(0, 10),
      });
      n++;
    }
  } else if (groupBy === "month") {
    const [y0, m0] = dateFrom.split("-").map(Number);
    const [y1, m1] = dateTo.split("-").map(Number);
    let y = y0, m = m0;
    const MNAMES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
    while (y < y1 || (y === y1 && m <= m1)) {
      buckets.push({ key: `${y}-${String(m).padStart(2, "0")}`, label: MNAMES[m - 1] });
      m++; if (m > 12) { m = 1; y++; }
    }
  }

  function bucketOf(date) {
    if (groupBy === "day") return date;
    if (groupBy === "month") return date.slice(0, 7);
    if (groupBy === "week") {
      const b = buckets.find((b) => date >= b.startISO && date <= b.endISO);
      return b ? b.key : null;
    }
    return null;
  }

  // ===== gráfica de pastel: gastos por categoría =====
  if (kind === "pie" || metric === "expense_by_category") {
    const byCat = {};
    pool.filter((t) => t.type === "expense" && t.categoryId).forEach((t) => {
      byCat[t.categoryId] = (byCat[t.categoryId] || 0) + t.amount;
    });
    const segments = Object.entries(byCat)
      .map(([id, value]) => {
        const c = config.categories.find((x) => x.id === id);
        return { label: c ? `${c.emoji} ${c.name}` : "Sin categoría", value };
      })
      .sort((a, b) => b.value - a.value);
    return { kind: "pie", title: spec.title || "Gastos por categoría", segments };
  }

  // ===== series temporales (line/bars) =====
  if (groupBy === "category") {
    // barras por categoría
    const byCat = {};
    pool.filter((t) => t.type === (metric === "income" ? "income" : "expense") && t.categoryId).forEach((t) => {
      byCat[t.categoryId] = (byCat[t.categoryId] || 0) + t.amount;
    });
    const sorted = Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const xLabels = sorted.map(([id]) => {
      const c = config.categories.find((x) => x.id === id);
      return c ? c.name : "—";
    });
    const values = sorted.map(([, v]) => v);
    return {
      kind: kind === "line" ? "bars" : kind, // forzar barras para categorías
      title: spec.title || (metric === "income" ? "Ingresos por categoría" : "Gastos por categoría"),
      xLabels,
      series: [{ name: metric === "income" ? "Ingresos" : "Gastos", color: metric === "income" ? "#2E6F4E" : "#B8482A", values }],
    };
  }

  // por tiempo
  const incValues = buckets.map(() => 0);
  const expValues = buckets.map(() => 0);
  const netValues = buckets.map(() => 0);
  for (const t of pool) {
    const bk = bucketOf(t.date);
    if (!bk) continue;
    const idx = buckets.findIndex((b) => b.key === bk);
    if (idx < 0) continue;
    if (t.type === "income") { incValues[idx] += t.amount; netValues[idx] += t.amount; }
    else { expValues[idx] += t.amount; netValues[idx] -= t.amount; }
  }

  const xLabels = buckets.map((b) => b.label);
  let series = [];
  if (metric === "income") series = [{ name: "Ingresos", color: "#2E6F4E", values: incValues }];
  else if (metric === "expense") series = [{ name: "Gastos", color: "#B8482A", values: expValues }];
  else if (metric === "net") series = [{ name: "Flujo neto", color: "#B0863A", values: netValues }];
  else series = [
    { name: "Ingresos", color: "#2E6F4E", values: incValues },
    { name: "Gastos", color: "#B8482A", values: expValues },
  ];

  const titleMap = {
    day: "por día", week: "por semana", month: "por mes",
  };
  const defaultTitle = spec.title
    || (metric === "net" ? `Flujo neto ${titleMap[groupBy] || ""}` : `Ingresos vs gastos ${titleMap[groupBy] || ""}`);

  return { kind, title: defaultTitle.trim(), xLabels, series };
}

/* buscar categoría / cuenta por id o por nombre */
function findCat(config, ref, accountId) {
  if (ref == null) return null;
  const r = norm(String(ref)).trim();
  const byId = config.categories.find((c) => c.id === ref);
  if (byId) return byId;
  const pool = accountId ? config.categories.filter((c) => c.accountId === accountId) : config.categories;
  return pool.find((c) => norm(c.name).trim() === r)
    || pool.find((c) => norm(c.name).includes(r) && r.length >= 3) || null;
}
function findAcc(config, ref) {
  if (ref == null) return null;
  const r = norm(String(ref)).trim();
  return config.accounts.find((a) => a.id === ref)
    || config.accounts.find((a) => norm(a.name).trim() === r)
    || config.accounts.find((a) => norm(a.name).includes(r) && r.length >= 3) || null;
}

/* ========================== RECURRENCIAS ================================
   config.recurring = [
     {
       id, type:"income"|"expense", amount, description,
       categoryId, accountId,
       freq:"daily"|"weekly"|"biweekly"|"monthly"|"yearly",
       dayOfMonth?:number (para monthly/yearly, 1-31),
       month?:number (para yearly, 1-12),
       daysOfWeek?:[0..6] (para weekly: 0=domingo),
       startDate:"YYYY-MM-DD",
       endDate?:"YYYY-MM-DD",
       paused?:boolean,
       lastRunDate?:"YYYY-MM-DD" (última fecha cubierta — incluida)
     }
   ]
   Cada movimiento generado por una recurrencia tiene t.fromRecurring = <id> */

function dateKeyDaysAfter(iso, n) {
  const d = new Date(iso + "T12:00:00");
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}

/* todas las fechas en las que la recurrencia debe ejecutarse desde "fromDate" hasta "toDate" inclusivo */
function recurringDatesBetween(rule, fromDate, toDate) {
  const dates = [];
  if (!rule || !rule.startDate) return dates;
  const start = rule.startDate;
  const end = rule.endDate || toDate;
  // Empezar desde el día siguiente a lastRunDate, o desde startDate si nunca corrió
  let cursor = rule.lastRunDate ? dateKeyDaysAfter(rule.lastRunDate, 1) : start;
  if (cursor < start) cursor = start;
  const realEnd = end < toDate ? end : toDate;
  // safety: limita a 365 días por procesado para no colgarse
  let safety = 0;
  while (cursor <= realEnd && safety < 800) {
    const d = new Date(cursor + "T12:00:00");
    let matches = false;
    if (rule.freq === "daily") matches = true;
    else if (rule.freq === "weekly") {
      const dow = d.getDay();
      const days = Array.isArray(rule.daysOfWeek) && rule.daysOfWeek.length ? rule.daysOfWeek : [new Date(start + "T12:00:00").getDay()];
      matches = days.includes(dow);
    }
    else if (rule.freq === "biweekly") {
      const startD = new Date(start + "T12:00:00");
      const diff = Math.round((d - startD) / 86400000);
      matches = diff >= 0 && diff % 14 === 0;
    }
    else if (rule.freq === "monthly") {
      const dom = rule.dayOfMonth || new Date(start + "T12:00:00").getDate();
      // si dom > último día del mes, usar último día
      const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
      const targetDom = Math.min(dom, lastDay);
      matches = d.getDate() === targetDom;
    }
    else if (rule.freq === "yearly") {
      const startD = new Date(start + "T12:00:00");
      const mon = (rule.month || (startD.getMonth() + 1)) - 1;
      const dom = rule.dayOfMonth || startD.getDate();
      matches = d.getMonth() === mon && d.getDate() === dom;
    }
    if (matches) dates.push(cursor);
    cursor = dateKeyDaysAfter(cursor, 1);
    safety++;
  }
  return dates;
}

/* Procesar todas las recurrencias activas. Devuelve { config, txs, generated } */
function processRecurring(config, txs) {
  const rules = Array.isArray(config.recurring) ? config.recurring : [];
  if (!rules.length) return { config, txs, generated: 0 };
  const todayK = today();
  const newTxs = [];
  const updatedRules = rules.map((rule) => {
    if (rule.paused) return rule;
    if (!rule.amount || !rule.accountId) return rule;
    // si la categoría ya no existe o no pertenece a la cuenta, dejar null
    let categoryId = rule.categoryId;
    const cat = config.categories.find((c) => c.id === categoryId);
    if (!cat || cat.accountId !== rule.accountId) categoryId = null;

    const fires = recurringDatesBetween(rule, rule.startDate, todayK);
    if (!fires.length) return rule;
    fires.forEach((date) => {
      newTxs.push({
        id: uid(),
        type: rule.type || "expense",
        amount: Number(rule.amount) || 0,
        description: rule.description || "(recurrencia)",
        categoryId,
        accountId: rule.accountId,
        date,
        fromRecurring: rule.id,
      });
    });
    return { ...rule, lastRunDate: fires[fires.length - 1] };
  });
  if (!newTxs.length) return { config, txs, generated: 0 };
  return {
    config: { ...config, recurring: updatedRules },
    txs: [...newTxs, ...txs],
    generated: newTxs.length,
  };
}

/* aplicar las acciones que devuelve el asistente */
function applyActions(config0, txs0, actions) {
  let config = { ...config0, categories: [...config0.categories], accounts: [...config0.accounts] };
  let txs = [...txs0];
  const log = [];
  const ok = (text) => log.push({ ok: true, text });
  const no = (text) => log.push({ ok: false, text });

  for (const a of Array.isArray(actions) ? actions : []) {
    if (!a || !a.type) continue;
    try {
      if (a.type === "add_category") {
        const ac = findAcc(config, a.accountId ?? a.accountName) || config.accounts[0];
        const seed = SEED_KW[norm(a.name || "").trim()];
        config = { ...config, categories: [...config.categories, {
          id: uid(), name: a.name || "Categoría", emoji: a.emoji || "📦",
          type: a.categoryType === "income" ? "income" : "expense",
          accountId: ac ? ac.id : null, keywords: seed ? [...seed] : [],
        }] };
        ok(`Categoría creada en ${ac ? ac.name : "?"}: ${a.emoji || "📦"} ${a.name || ""}`);

      } else if (a.type === "remove_category") {
        const c = findCat(config, a.id ?? a.name, findAcc(config, a.accountId ?? a.accountName)?.id);
        if (c) { config = { ...config, categories: config.categories.filter((x) => x.id !== c.id) }; ok(`Categoría eliminada: ${c.name}`); }
        else no(`No encontré la categoría «${a.name ?? a.id}»`);

      } else if (a.type === "edit_category") {
        const c = findCat(config, a.id ?? a.name, findAcc(config, a.accountId ?? a.accountName)?.id);
        if (c) {
          const u = { ...c };
          if (a.newName) u.name = a.newName;
          if (a.newEmoji) u.emoji = a.newEmoji;
          if (a.newType) u.type = a.newType === "income" ? "income" : "expense";
          config = { ...config, categories: config.categories.map((x) => x.id === c.id ? u : x) };
          ok(`Categoría actualizada: ${u.emoji} ${u.name}`);
        } else no(`No encontré la categoría «${a.name ?? a.id}»`);

      } else if (a.type === "set_account_categories") {
        const ac = findAcc(config, a.accountId ?? a.accountName);
        const incoming = Array.isArray(a.categories) ? a.categories : [];
        if (!ac) no(`No encontré la cuenta «${a.accountName ?? a.accountId}»`);
        else if (!incoming.length) no(`No me diste las categorías para «${ac.name}»`);
        else {
          const ofAcc = config.categories.filter((c) => c.accountId === ac.id);
          const oldExpense = ofAcc.filter((c) => c.type === "expense");
          const keepIncome = ofAcc.filter((c) => c.type === "income");
          const newExpense = incoming.map((it) => {
            const nm = (it.name || "Categoría").trim();
            const prev = oldExpense.find((c) => norm(c.name).trim() === norm(nm).trim());
            const seed = SEED_KW[norm(nm).trim()];
            return {
              id: uid(), name: nm, emoji: it.emoji || (prev ? prev.emoji : "📦"),
              type: "expense", accountId: ac.id,
              keywords: prev ? [...(prev.keywords || [])] : (seed ? [...seed] : []),
            };
          });
          // remapear movimientos por nombre para no perder historial
          const nameToNew = {};
          newExpense.forEach((c) => { nameToNew[norm(c.name).trim()] = c.id; });
          txs = txs.map((t) => {
            const oc = oldExpense.find((c) => c.id === t.categoryId);
            if (!oc) return t;
            const nid = nameToNew[norm(oc.name).trim()];
            return { ...t, categoryId: nid || null };
          });
          const others = config.categories.filter((c) => c.accountId !== ac.id);
          config = { ...config, categories: [...others, ...keepIncome, ...newExpense] };
          ok(`Categorías de gasto de ${ac.name} actualizadas (${newExpense.length})`);
        }

      } else if (a.type === "add_account") {
        const nid = uid();
        config = {
          ...config,
          accounts: [...config.accounts, { id: nid, name: a.name || "Cuenta", initialBalance: Number(a.initialBalance) || 0 }],
          categories: [...config.categories, ...defaultCatsForAccount(nid)],
          accountMode: "multiple",
        };
        ok(`Cuenta creada: ${a.name || ""} (saldo inicial ${fmt(Number(a.initialBalance) || 0)})`);

      } else if (a.type === "remove_account") {
        const ac = findAcc(config, a.id ?? a.name);
        if (!ac) no(`No encontré la cuenta «${a.name ?? a.id}»`);
        else if (config.accounts.length <= 1) no(`No puedo eliminar la única cuenta.`);
        else {
          config = {
            ...config,
            accounts: config.accounts.filter((x) => x.id !== ac.id),
            categories: config.categories.filter((c) => c.accountId !== ac.id),
          };
          config.accountMode = config.accounts.length > 1 ? "multiple" : "single";
          ok(`Cuenta eliminada: ${ac.name}`);
        }

      } else if (a.type === "edit_account") {
        const ac = findAcc(config, a.id ?? a.name);
        if (ac) {
          const u = { ...ac };
          if (a.newName) u.name = a.newName;
          if (a.newInitialBalance != null) u.initialBalance = Number(a.newInitialBalance) || 0;
          config = { ...config, accounts: config.accounts.map((x) => x.id === ac.id ? u : x) };
          ok(`Cuenta actualizada: ${u.name}`);
        } else no(`No encontré la cuenta «${a.name ?? a.id}»`);

      } else if (a.type === "add_transaction") {
        const ac = findAcc(config, a.accountName) || config.accounts[0];
        const c = findCat(config, a.categoryName, ac ? ac.id : null);
        const tx = {
          id: uid(), type: a.txType === "income" ? "income" : "expense",
          amount: Math.abs(Number(a.amount) || 0), description: a.description || "",
          categoryId: c ? c.id : null, accountId: ac ? ac.id : null, date: a.date || today(),
          passThrough: !!a.passThrough, linkedTxId: a.linkedTxId || null,
        };
        txs = [tx, ...txs];
        ok(`${tx.type === "income" ? "Ingreso" : "Gasto"}${tx.passThrough ? " (de paso)" : ""} registrado: ${tx.description || (c ? c.name : "")} · ${fmt(tx.amount)}`);

      } else if (a.type === "delete_transaction") {
        const t = txs.find((x) => x.id === a.id);
        if (t) { txs = txs.filter((x) => x.id !== a.id); ok(`Movimiento eliminado: ${t.description || fmt(t.amount)}`); }
        else no(`No encontré ese movimiento`);

      } else if (a.type === "edit_transaction") {
        const t = txs.find((x) => x.id === a.id);
        if (t) {
          const u = { ...t };
          if (a.amount != null) u.amount = Math.abs(Number(a.amount) || 0);
          if (a.description != null) u.description = a.description;
          if (a.txType) u.type = a.txType === "income" ? "income" : "expense";
          if (a.date) u.date = a.date;
          if (a.accountName) { const ac = findAcc(config, a.accountName); if (ac) u.accountId = ac.id; }
          if (a.categoryName) { const c = findCat(config, a.categoryName, u.accountId); if (c) u.categoryId = c.id; }
          if (a.passThrough != null) u.passThrough = !!a.passThrough;
          if (a.linkedTxId !== undefined) u.linkedTxId = a.linkedTxId || null;
          txs = txs.map((x) => x.id === t.id ? u : x);
          ok(`Movimiento actualizado: ${u.description || fmt(u.amount)}`);
        } else no(`No encontré ese movimiento`);

      } else if (a.type === "mark_passthrough") {
        // marca uno o dos movimientos como "no real" y opcionalmente los vincula
        const ids = Array.isArray(a.ids) ? a.ids : (a.id ? [a.id] : []);
        const found = ids.map((id) => txs.find((x) => x.id === id)).filter(Boolean);
        if (!found.length) { no(`No encontré los movimientos a marcar como dinero de paso`); }
        else {
          // si son exactamente 2 y son de tipos opuestos, vincularlos
          let linkA = null, linkB = null;
          if (found.length === 2 && found[0].type !== found[1].type) {
            linkA = found[0].id; linkB = found[1].id;
          }
          const updates = new Map();
          found.forEach((t) => {
            const u = { ...t, passThrough: true };
            if (linkA && linkB) u.linkedTxId = t.id === linkA ? linkB : linkA;
            updates.set(t.id, u);
          });
          txs = txs.map((x) => updates.has(x.id) ? updates.get(x.id) : x);
          if (linkA && linkB) {
            const inc = found.find((t) => t.type === "income");
            const exp = found.find((t) => t.type === "expense");
            const diff = (inc?.amount || 0) - (exp?.amount || 0);
            const note = Math.abs(diff) < 0.01 ? "Sin diferencia neta"
              : diff > 0 ? `Ganancia neta: ${fmt(diff)}` : `Pusiste de tu bolsa: ${fmt(Math.abs(diff))}`;
            ok(`Vinculados como reembolso: ${inc?.description || fmt(inc?.amount || 0)} ↔ ${exp?.description || fmt(exp?.amount || 0)}. ${note}`);
          } else {
            found.forEach((t) => ok(`Marcado como dinero de paso: ${t.description || fmt(t.amount)}`));
          }
        }

      } else if (a.type === "unmark_passthrough") {
        const ids = Array.isArray(a.ids) ? a.ids : (a.id ? [a.id] : []);
        const targets = ids.map((id) => txs.find((x) => x.id === id)).filter(Boolean);
        if (!targets.length) { no(`No encontré los movimientos a desmarcar`); }
        else {
          const setIds = new Set();
          targets.forEach((t) => {
            setIds.add(t.id);
            if (t.linkedTxId) setIds.add(t.linkedTxId);
          });
          txs = txs.map((x) => setIds.has(x.id) ? { ...x, passThrough: false, linkedTxId: null } : x);
          ok(`Desmarcados ${setIds.size} movimiento${setIds.size === 1 ? "" : "s"} (vuelven a contar como reales)`);
        }

      /* ============== RECURRENCIAS ============== */
      } else if (a.type === "add_recurring") {
        const ac = findAcc(config, a.accountId ?? a.accountName);
        if (!ac) { no(`No encontré la cuenta «${a.accountName ?? a.accountId}»`); }
        else {
          const cat = a.categoryName || a.categoryId
            ? findCat(config, a.categoryId ?? a.categoryName, ac.id)
            : null;
          const validFreqs = ["daily","weekly","biweekly","monthly","yearly"];
          const freq = validFreqs.includes(a.freq) ? a.freq : "monthly";
          const rule = {
            id: uid(),
            type: a.txType === "income" ? "income" : "expense",
            amount: Number(a.amount) || 0,
            description: a.description || "",
            categoryId: cat ? cat.id : null,
            accountId: ac.id,
            freq,
            dayOfMonth: a.dayOfMonth != null ? Number(a.dayOfMonth) : undefined,
            month: a.month != null ? Number(a.month) : undefined,
            daysOfWeek: Array.isArray(a.daysOfWeek) ? a.daysOfWeek.map(Number) : undefined,
            startDate: a.startDate || today(),
            endDate: a.endDate || undefined,
            paused: false,
            active: true,
          };
          config = { ...config, recurring: [...(config.recurring || []), rule] };
          // ejecutar de inmediato lo pendiente
          const r = processRecurring(config, txs);
          config = r.config; txs = r.txs;
          const freqLabel = { daily: "diaria", weekly: "semanal", biweekly: "quincenal", monthly: "mensual", yearly: "anual" }[freq];
          ok(`Recurrencia ${freqLabel} creada: ${rule.type === "income" ? "+" : "−"}$${rule.amount} en ${ac.name}${r.generated > 0 ? ` (se aplicaron ${r.generated} de inmediato)` : ""}`);
        }

      } else if (a.type === "edit_recurring") {
        const rules = config.recurring || [];
        const rule = rules.find((r) => r.id === a.id);
        if (!rule) { no(`No encontré la recurrencia con id ${a.id}`); }
        else {
          const u = { ...rule };
          if (a.amount != null) u.amount = Number(a.amount);
          if (a.description != null) u.description = a.description;
          if (a.txType) u.type = a.txType === "income" ? "income" : "expense";
          if (a.freq) u.freq = a.freq;
          if (a.dayOfMonth != null) u.dayOfMonth = Number(a.dayOfMonth);
          if (a.month != null) u.month = Number(a.month);
          if (Array.isArray(a.daysOfWeek)) u.daysOfWeek = a.daysOfWeek.map(Number);
          if (a.startDate) u.startDate = a.startDate;
          if (a.endDate !== undefined) u.endDate = a.endDate || undefined;
          if (a.accountId || a.accountName) {
            const ac = findAcc(config, a.accountId ?? a.accountName);
            if (ac) u.accountId = ac.id;
          }
          if (a.categoryId || a.categoryName) {
            const cat = findCat(config, a.categoryId ?? a.categoryName, u.accountId);
            if (cat) u.categoryId = cat.id;
          }
          config = { ...config, recurring: rules.map((r) => r.id === u.id ? u : r) };
          ok(`Recurrencia actualizada`);
        }

      } else if (a.type === "delete_recurring") {
        const rules = config.recurring || [];
        const before = rules.length;
        const newRules = rules.filter((r) => r.id !== a.id);
        if (newRules.length === before) { no(`No encontré esa recurrencia`); }
        else {
          config = { ...config, recurring: newRules };
          // borrar también los movimientos ya generados, si el usuario lo pide explícito
          if (a.alsoDeleteGenerated) {
            const removed = txs.filter((t) => t.fromRecurring === a.id).length;
            txs = txs.filter((t) => t.fromRecurring !== a.id);
            ok(`Recurrencia eliminada y ${removed} movimientos generados también borrados`);
          } else {
            ok(`Recurrencia eliminada (los movimientos ya generados se mantienen)`);
          }
        }

      } else if (a.type === "pause_recurring") {
        const rules = config.recurring || [];
        const rule = rules.find((r) => r.id === a.id);
        if (!rule) { no(`No encontré esa recurrencia`); }
        else {
          config = { ...config, recurring: rules.map((r) => r.id === a.id ? { ...r, paused: true } : r) };
          ok(`Recurrencia pausada`);
        }

      } else if (a.type === "resume_recurring") {
        const rules = config.recurring || [];
        const rule = rules.find((r) => r.id === a.id);
        if (!rule) { no(`No encontré esa recurrencia`); }
        else {
          config = { ...config, recurring: rules.map((r) => r.id === a.id ? { ...r, paused: false } : r) };
          // procesar lo pendiente
          const r2 = processRecurring(config, txs);
          config = r2.config; txs = r2.txs;
          ok(`Recurrencia reanudada${r2.generated > 0 ? ` (se aplicaron ${r2.generated} movimientos pendientes)` : ""}`);
        }

      /* ============== ACCIONES MASIVAS (BULK) ============== */
      } else if (a.type === "bulk_edit_category") {
        // re-categoriza muchos movimientos. Filtros opcionales: ids, fromCategoryId, accountId, keyword, fromDate, toDate
        const ac = a.accountId || a.accountName ? findAcc(config, a.accountId ?? a.accountName) : null;
        const targetCat = findCat(config, a.toCategoryId ?? a.toCategoryName, ac?.id);
        if (!targetCat) { no(`No encontré la categoría destino «${a.toCategoryName ?? a.toCategoryId}»`); }
        else {
          const fromCat = a.fromCategoryId || a.fromCategoryName ? findCat(config, a.fromCategoryId ?? a.fromCategoryName, ac?.id) : null;
          const ids = Array.isArray(a.ids) ? new Set(a.ids) : null;
          const kw = a.keyword ? norm(a.keyword) : null;
          let touched = 0;
          txs = txs.map((t) => {
            if (ids && !ids.has(t.id)) return t;
            if (!ids && ac && t.accountId !== ac.id) return t;
            if (!ids && fromCat && t.categoryId !== fromCat.id) return t;
            if (!ids && kw && !norm(t.description || "").includes(kw)) return t;
            if (!ids && a.fromDate && t.date < a.fromDate) return t;
            if (!ids && a.toDate && t.date > a.toDate) return t;
            if (t.categoryId === targetCat.id) return t;
            touched++;
            return { ...t, categoryId: targetCat.id };
          });
          if (touched > 0) ok(`Re-categorizados ${touched} movimientos a ${targetCat.emoji} ${targetCat.name}`);
          else no(`No encontré movimientos que coincidan con los filtros`);
        }

      } else if (a.type === "bulk_delete") {
        const ids = Array.isArray(a.ids) ? new Set(a.ids) : null;
        const ac = a.accountId || a.accountName ? findAcc(config, a.accountId ?? a.accountName) : null;
        const fromCat = a.categoryId || a.categoryName ? findCat(config, a.categoryId ?? a.categoryName, ac?.id) : null;
        const kw = a.keyword ? norm(a.keyword) : null;
        const before = txs.length;
        txs = txs.filter((t) => {
          if (ids) return !ids.has(t.id);
          // criterios combinables
          let match = true;
          if (ac && t.accountId !== ac.id) match = false;
          if (fromCat && t.categoryId !== fromCat.id) match = false;
          if (kw && !norm(t.description || "").includes(kw)) match = false;
          if (a.fromDate && t.date < a.fromDate) match = false;
          if (a.toDate && t.date > a.toDate) match = false;
          if (a.txType && t.type !== a.txType) match = false;
          return !match;  // si NO matchea todos los filtros, se conserva
        });
        const removed = before - txs.length;
        if (removed > 0) ok(`Eliminados ${removed} movimientos`);
        else no(`No encontré movimientos que coincidan con los filtros`);

      } else if (a.type === "bulk_move_account") {
        const target = findAcc(config, a.toAccountId ?? a.toAccountName);
        if (!target) { no(`No encontré la cuenta destino`); }
        else {
          const ids = Array.isArray(a.ids) ? new Set(a.ids) : null;
          const fromAc = a.fromAccountId || a.fromAccountName ? findAcc(config, a.fromAccountId ?? a.fromAccountName) : null;
          const kw = a.keyword ? norm(a.keyword) : null;
          let touched = 0;
          txs = txs.map((t) => {
            let match = true;
            if (ids) match = ids.has(t.id);
            else {
              if (fromAc && t.accountId !== fromAc.id) match = false;
              if (kw && !norm(t.description || "").includes(kw)) match = false;
              if (a.fromDate && t.date < a.fromDate) match = false;
              if (a.toDate && t.date > a.toDate) match = false;
            }
            if (!match) return t;
            if (t.accountId === target.id) return t;
            touched++;
            // al mover de cuenta, la categoría ya no aplica (pertenece a otra cuenta)
            return { ...t, accountId: target.id, categoryId: null };
          });
          if (touched > 0) ok(`Movidos ${touched} movimientos a ${target.name} (su categoría se reinició)`);
          else no(`No encontré movimientos que coincidan`);
        }

      /* ============== CONFIGURACIÓN ============== */
      } else if (a.type === "set_date_range") {
        const validPresets = ["today","week","month","last-month","3m","6m","year","last-year","all","custom"];
        const preset = validPresets.includes(a.preset) ? a.preset : "month";
        const newRange = preset === "custom"
          ? { preset: "custom", from: a.from || today(), to: a.to || today() }
          : { preset };
        config = { ...config, dateRange: newRange };
        ok(`Rango cambiado a: ${rangeLabel(newRange)}`);

      } else {
        no(`Acción no reconocida: ${a.type}`);
      }
    } catch (e) {
      no(`Hubo un error aplicando un cambio.`);
    }
  }
  return { config, txs, log };
}

/* system prompt del asistente, con el estado actual de la app */
function assistantSystem(config, txs) {
  const cuentas = config.accounts.map((a) => ({
    id: a.id, nombre: a.name, saldoInicial: a.initialBalance || 0, saldoActual: accountBalance(config, txs, a.id),
    categorias: config.categories.filter((c) => c.accountId === a.id)
      .map((c) => ({ id: c.id, nombre: c.name, emoji: c.emoji, tipo: c.type })),
  }));
  // hasta 100 movimientos recientes (de los más nuevos)
  const sorted = [...txs].sort((a, b) => b.date.localeCompare(a.date) || (b.id || "").localeCompare(a.id || ""));
  const recent = sorted.slice(0, 100).map((t) => ({
    id: t.id, tipo: t.type, monto: t.amount, concepto: t.description,
    categoria: (config.categories.find((c) => c.id === t.categoryId) || {}).name || null,
    cuenta: (config.accounts.find((a) => a.id === t.accountId) || {}).name || null, fecha: t.date,
    dineroDePaso: !!t.passThrough,
    vinculadoCon: t.linkedTxId || null,
    deRecurrencia: t.fromRecurring || null,
  }));

  /* ============ AGREGADOS PRE-CALCULADOS ============ */
  // rango activo del usuario
  const range = config.dateRange || DEFAULT_RANGE;
  const { from: rangeFrom, to: rangeTo } = resolveRange(range);
  const rangeTxs = txsInRange(txs, range);
  const rangeStat = statTxs(rangeTxs).all;
  const rangeIncome = rangeStat.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const rangeExpense = rangeStat.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  // top categorías de gasto e ingreso en el rango
  const expByCat = {}, incByCat = {};
  rangeStat.forEach((t) => {
    if (!t.categoryId) return;
    if (t.type === "expense") expByCat[t.categoryId] = (expByCat[t.categoryId] || 0) + t.amount;
    else incByCat[t.categoryId] = (incByCat[t.categoryId] || 0) + t.amount;
  });
  const topRows = (obj) => Object.entries(obj)
    .map(([id, amt]) => {
      const c = config.categories.find((x) => x.id === id);
      return c ? { categoria: c.name, monto: amt } : null;
    }).filter(Boolean).sort((a, b) => b.monto - a.monto).slice(0, 8);

  // saldo neto histórico
  const saldoNeto = config.accounts.reduce((s, a) => s + accountBalance(config, txs, a.id), 0);

  // por mes (últimos 6 meses) para comparaciones
  const monthly = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date();
    d.setMonth(d.getMonth() - i);
    const mk = d.toISOString().slice(0, 7);
    const monthTxs = txs.filter((t) => t.date.slice(0, 7) === mk);
    const ms = statTxs(monthTxs).all;
    monthly.push({
      mes: mk,
      ingresos: ms.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0),
      gastos: ms.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0),
      movimientos: monthTxs.length,
    });
  }

  // patrones frecuentes detectados (últimos 2 meses)
  const patterns = detectFrequentPatterns(txs, config);
  const patronesFrecuentes = patterns.map(p => ({
    concepto: p.description,
    monto: p.amount,
    tipo: p.type,
    categoria: p.categoryName,
    veces: p.count,
    ultimaFecha: p.lastDate,
    esAnual: !!p.annual,
    esRecurrente: !!(config.recurring || []).find(r =>
      r.description.toLowerCase().trim() === p.description.toLowerCase().trim()),
  }));

  const agregados = {
    rango: { preset: range.preset, desde: rangeFrom, hasta: rangeTo, etiqueta: rangeLabel(range) },
    enRango: { ingresos: rangeIncome, gastos: rangeExpense, flujoNeto: rangeIncome - rangeExpense, movimientos: rangeTxs.length },
    topGastosEnRango: topRows(expByCat),
    topIngresosEnRango: topRows(incByCat),
    porMes6m: monthly,
    saldoNetoGlobal: saldoNeto,
    patronesFrecuentes,
  };

  // recurrencias actuales
  const recurrencias = (config.recurring || []).map((r) => ({
    id: r.id,
    tipo: r.type,
    monto: r.amount,
    concepto: r.description,
    cuenta: (config.accounts.find((a) => a.id === r.accountId) || {}).name || null,
    categoria: (config.categories.find((c) => c.id === r.categoryId) || {}).name || null,
    frecuencia: r.freq,
    diaDelMes: r.dayOfMonth || null,
    diasDeSemana: r.daysOfWeek || null,
    desde: r.startDate,
    hasta: r.endDate || null,
    pausada: !!r.paused,
    ultimaCorrida: r.lastRunDate || null,
  }));

  return `Eres el asistente de "Zafi", una app de finanzas personales con IA. El usuario habla español de México. Te pide cosas en lenguaje natural y tú las ejecutas o respondes sus dudas.

IMPORTANTE: cada CUENTA tiene SU PROPIA lista de categorías. Las categorías NO son globales; pertenecen a una cuenta.

ESTADO ACTUAL DE LA APP:
Cuentas (cada una con sus categorías): ${JSON.stringify(cuentas)}
Movimientos recientes (últimos 100): ${JSON.stringify(recent)}
Recurrencias activas: ${JSON.stringify(recurrencias)}
Agregados pre-calculados: ${JSON.stringify(agregados)}
Fecha de hoy: ${today()}

PATRONES FRECUENTES (últimos 2 meses):
${patronesFrecuentes.length > 0 ? JSON.stringify(patronesFrecuentes) : "Ninguno detectado aún."}
Si hay patrones frecuentes que NO son recurrentes aún (esRecurrente=false), puedes sugerirle al usuario crear una recurrencia para automatizar esos movimientos. Si esAnual=true, el patrón se repite también el año pasado (mismo periodo). Ejemplo: "Noté que registras 'Spotify' cada mes por $189. ¿Quieres que lo haga recurrente para que se registre solo?" o "El año pasado también pagaste 'Seguro auto' por esta fecha por $12,000 — ¿quieres que te recuerde?" Pero NO insistas si el usuario no quiere.

RESPONDE SIEMPRE con UN SOLO objeto JSON válido, sin markdown ni texto fuera del JSON:
{"message":"texto breve para el usuario","actions":[ ... ]}

ACCIONES disponibles (usa "actions":[] si solo respondes una pregunta):
- {"type":"set_account_categories","accountId":"<id de la cuenta>","categories":[{"name":"...","emoji":"<emoji>"},...]}  → REEMPLAZA todas las categorías de GASTO de esa cuenta por la lista dada. Úsala cuando el usuario pida "cambia/pon las categorías de <cuenta> a esto: ...". Ponle un emoji apropiado a cada una. Las categorías de ingreso de esa cuenta NO se tocan.
- {"type":"add_category","accountId":"<id de la cuenta>","name":"...","emoji":"<emoji>","categoryType":"income"|"expense"}
- {"type":"edit_category","id":"<id de la categoría>","newName":"...","newEmoji":"...","newType":"income"|"expense"}
- {"type":"remove_category","id":"<id de la categoría>"}
- {"type":"add_account","name":"...","initialBalance":<número>}
- {"type":"edit_account","id":"<id de la cuenta>","newName":"...","newInitialBalance":<número>}
- {"type":"remove_account","id":"<id de la cuenta>"}
- {"type":"add_transaction","txType":"income"|"expense","amount":<número>,"description":"...","categoryName":"<categoría existente de esa cuenta>","accountName":"<nombre de cuenta>","date":"YYYY-MM-DD","passThrough":<true|false opcional>}
- {"type":"edit_transaction","id":"<id del movimiento>","amount":<número>,"description":"...","categoryName":"...","accountName":"...","txType":"...","date":"...","passThrough":<true|false>,"linkedTxId":"<id o null>"}
- {"type":"delete_transaction","id":"<id del movimiento>"}
- {"type":"show_chart","spec":{ ... }}  → muestra una gráfica inline en el chat. La spec define qué graficar (ver sección 📊 abajo). Úsala cuando el usuario te PIDA una gráfica.

🔁 RECURRENCIAS (movimientos automáticos):
- {"type":"add_recurring","txType":"income"|"expense","amount":<n>,"description":"...","accountName":"...","categoryName":"...","freq":"daily"|"weekly"|"biweekly"|"monthly"|"yearly","dayOfMonth":<1-31 opc>,"month":<1-12 opc>,"daysOfWeek":[<0-6> opc],"startDate":"YYYY-MM-DD opc","endDate":"YYYY-MM-DD opc"}
   Ejemplos para freq:
   • "daily" → todos los días desde startDate (default: hoy)
   • "weekly" + daysOfWeek:[1] → cada lunes (0=domingo, 1=lunes, ..., 6=sábado)
   • "biweekly" → cada 14 días desde startDate
   • "monthly" + dayOfMonth:15 → el día 15 de cada mes (si el mes no llega a ese día, se aplica el último)
   • "yearly" + month:12 + dayOfMonth:25 → cada 25 de diciembre
   Cuando creas una recurrencia, los movimientos pendientes desde startDate hasta hoy se aplican automáticamente.
- {"type":"edit_recurring","id":"<id>","amount":...,"description":"...","freq":"...","dayOfMonth":...,"endDate":"..."} — solo incluye los campos que cambian
- {"type":"delete_recurring","id":"<id>","alsoDeleteGenerated":<true|false opc>}  → si alsoDeleteGenerated=true, también borra todos los movimientos ya creados por esa recurrencia
- {"type":"pause_recurring","id":"<id>"}  → la regla deja de ejecutarse hasta que la reanuden
- {"type":"resume_recurring","id":"<id>"}  → al reanudar, aplica todos los movimientos pendientes desde lastRunDate hasta hoy

📦 ACCIONES MASIVAS (BULK):
- {"type":"bulk_edit_category","toCategoryName":"<categoría destino>","accountName":"<cuenta opc>","fromCategoryName":"<cuenta origen opc>","keyword":"<filtro opc>","fromDate":"YYYY-MM-DD opc","toDate":"YYYY-MM-DD opc","ids":["<id1>",...] (opc)}
   Re-categoriza varios movimientos al mismo tiempo. Filtra por ids, o por cualquier combinación de: accountName + fromCategoryName + keyword (en concepto) + rango de fechas.
   Ejemplos: "marca todo lo que diga OXXO como categoría Tiendas" → keyword:"OXXO", toCategoryName:"Tiendas".
- {"type":"bulk_delete","ids":["<id1>",...] (opc),"accountName":"...","categoryName":"...","keyword":"...","txType":"income"|"expense","fromDate":"...","toDate":"..."}
   Elimina varios movimientos. Filtra por ids, o por combinación de filtros (todos opcionales pero AL MENOS uno requerido).
   Ejemplos: "borra todos los Spotify" → keyword:"Spotify". "borra los gastos del 15 de mayo" → fromDate:"2026-05-15", toDate:"2026-05-15", txType:"expense".
- {"type":"bulk_move_account","toAccountName":"<cuenta destino>","fromAccountName":"<opc>","keyword":"<opc>","fromDate":"...","toDate":"...","ids":[...]}
   Mueve movimientos a otra cuenta. NOTA: al cambiar de cuenta la categoría se reinicia (porque pertenecía a otra cuenta).

⚙️ CONFIGURACIÓN:
- {"type":"set_date_range","preset":"today"|"week"|"month"|"last-month"|"3m"|"6m"|"year"|"last-year"|"all"|"custom","from":"YYYY-MM-DD (custom)","to":"YYYY-MM-DD (custom)"}
   Cambia el rango global de toda la app (afecta Inicio, Movimientos, Categorías y Estadísticas).

REGLAS:
- Para editar/eliminar categorías, cuentas o movimientos usa SIEMPRE el "id" exacto del estado de arriba.
- Para set_account_categories y add_category usa el "id" de la cuenta correcta.
- Si el usuario lista varias categorías para una cuenta, asume que son categorías de GASTO salvo que diga lo contrario.
- Si pide cambiar categorías de varias cuentas, devuelve una acción set_account_categories por cada cuenta.
- Para add_transaction, "categoryName" debe existir DENTRO de la cuenta indicada.
- Cada categoría nueva necesita un emoji apropiado.
- En edit_* incluye solo los campos que cambian.
- "message" debe ser corto y confirmar en español lo que hiciste.
- Si te piden algo que no puedes hacer, dilo en "message" sin inventar acciones.
- Si la petición es ambigua, pregunta en "message" y NO ejecutes acciones todavía.
- Para preguntas sobre gastos/ingresos usa los datos del estado de arriba.

SALDO INICIAL DE UNA CUENTA — IMPORTANTE:
El "saldoInicial" de cada cuenta es el monto base que tenía la cuenta antes de empezar a registrar movimientos. Es el campo "initialBalance" del modelo y se actualiza con la acción "edit_account" pasando "newInitialBalance".
Cuando el usuario diga cosas como:
- "ponle saldo inicial a Santander de 19039"
- "el saldo inicial de BBVA es 5000"
- "Santander empezó con 19039"
- "saldo inicial de Santander mayo 19039" (interpreta el mes como contexto, no como campo separado)
…SIEMPRE usa la acción {"type":"edit_account","id":"<id de la cuenta>","newInitialBalance":<número>}. Quita comas y signos del número: "19,039" → 19039.

🧠 INTELIGENCIA ANALÍTICA:
Tienes pre-calculado en "Agregados pre-calculados" todo lo necesario para responder preguntas analíticas SIN inventar números:
- agregados.enRango → ingresos, gastos, flujo neto del rango activo
- agregados.topGastosEnRango y topIngresosEnRango → categorías ordenadas con monto
- agregados.porMes6m → últimos 6 meses (cada uno con ingresos, gastos, movimientos) para comparaciones
- agregados.saldoNetoGlobal → patrimonio neto sumando todas las cuentas

Cuando el usuario pregunte cosas como:
- "¿cómo voy este mes?" → resume con agregados.enRango y top categorías
- "compara mayo con abril" → usa porMes6m, calcula la diferencia porcentual
- "¿en qué gasto más?" → top de topGastosEnRango
- "¿cuánto he gastado en Uber?" → busca en "Movimientos recientes" los que tienen "uber" en concepto y suma. Si necesitas mostrarlos, NO uses una acción — solo describe en "message". Si quieres mostrarlos visualmente, usa show_chart.
- "dame mis 10 gastos más grandes" → de "Movimientos recientes" filtra tipo expense, ordena por monto desc, toma los primeros 10, descríbelos
- "¿cuándo me toca el sueldo?" → revisa "Recurrencias activas" y calcula la próxima fecha

Siempre que respondas con números, redondea a pesos enteros (sin decimales) salvo que el usuario los pida explícitos. Formato: $1,234 con coma de miles.

CAPTURA MÚLTIPLE EN UN MENSAJE:
Si el usuario dice varios movimientos en una sola frase ("pagué 250 oxxo, recibí 5000, compré 1200 walmart"), devuelve UN array de varias acciones add_transaction. NO le preguntes uno por uno; intenta inferir todo y al final del "message" pregunta si algo quedó mal.

EJEMPLOS de respuesta correcta:

📊 GRÁFICAS — show_chart:
Cuando el usuario pida ver una gráfica, devuelve una acción show_chart con una "spec" que la app convierte en gráfica inline. NO calcules los datos tú, solo describe qué quieres graficar.

Formato de spec:
{
  "kind": "line" | "bars" | "pie",
  "title": "Texto del título (opcional)",
  "groupBy": "day" | "week" | "month" | "category" | "account",
  "metric": "income" | "expense" | "both" | "net" | "expense_by_category",
  "dateFrom": "YYYY-MM-DD",
  "dateTo": "YYYY-MM-DD",
  "accountId": "<id de cuenta o null para todas>",
  "includePassThrough": false
}

GUÍA PARA ELEGIR kind:
- "line" → tendencias en el tiempo (cómo evoluciona algo día/semana/mes)
- "bars" → comparar periodos o categorías lado a lado
- "pie" → distribución (gastos por categoría)

GUÍA PARA metric:
- "income" → solo ingresos
- "expense" → solo gastos
- "both" → dos series, ingresos y gastos por separado (lo más común para "ingresos vs gastos")
- "net" → flujo neto (ingresos − gastos por periodo)
- "expense_by_category" → gastos agrupados por categoría (se fuerza kind=pie)

GUÍA PARA fechas:
- "este mes" → dateFrom = primer día del mes actual, dateTo = hoy (${today()})
- "mayo 2026" → dateFrom = "2026-05-01", dateTo = "2026-05-31"
- "últimos 3 meses" → dateFrom = primer día de hace 3 meses, dateTo = hoy
- "este año" → dateFrom = "${today().slice(0, 4)}-01-01", dateTo = hoy
- "semana pasada" / "esta semana" → calcula el rango lunes-domingo correspondiente

GUÍA PARA groupBy:
- "day" → para rangos cortos (una semana o quincena)
- "week" → para 1 a 3 meses
- "month" → para 3 a 12 meses
- "category" → cuando preguntan "por categoría" sin pastel
- "account" → cuando comparan cuentas

REGLAS DE GRÁFICAS:
- Si el usuario menciona una sola cuenta ("solo de BBVA", "en Santander"), pon "accountId" con el id correcto.
- Si no especifica cuenta, deja accountId en null para usar todas.
- Por defecto, NO incluyas pass-through (includePassThrough: false) para que los reembolsos no inflen las cifras.
- En "message" describe brevemente qué muestra la gráfica.
- Puedes devolver más de una gráfica en una sola respuesta (varias show_chart en actions).

EJEMPLOS DE GRÁFICAS:

Usuario: "muéstrame una gráfica semana a semana del mes de mayo, ingresos vs gastos de líneas"
Tu respuesta:
{"message":"Aquí tienes mayo por semana 👇","actions":[{"type":"show_chart","spec":{"kind":"line","title":"Mayo · ingresos vs gastos","groupBy":"week","metric":"both","dateFrom":"${today().slice(0,4)}-05-01","dateTo":"${today().slice(0,4)}-05-31"}}]}

Usuario: "pastel de mis gastos del mes"
Tu respuesta:
{"message":"Tus gastos del mes por categoría:","actions":[{"type":"show_chart","spec":{"kind":"pie","metric":"expense_by_category","dateFrom":"${today().slice(0,7)}-01","dateTo":"${today()}"}}]}

Usuario: "barras de gastos por categoría este mes"
Tu respuesta:
{"message":"Aquí tienes:","actions":[{"type":"show_chart","spec":{"kind":"bars","groupBy":"category","metric":"expense","dateFrom":"${today().slice(0,7)}-01","dateTo":"${today()}"}}]}

Usuario: "cómo van mis ingresos los últimos 6 meses"
Tu respuesta (con cálculo de rango: hoy es ${today()}, hace 6 meses sería el primer día del mes que es 5 meses atrás):
{"message":"Tus ingresos de los últimos 6 meses:","actions":[{"type":"show_chart","spec":{"kind":"line","groupBy":"month","metric":"income","dateFrom":"<calcula>","dateTo":"${today()}"}}]}

Usuario: "compara BBVA y Santander este mes"
Cuando el usuario quiere comparar 2 cuentas, devuelve DOS gráficas, una por cuenta:
{"message":"Aquí los tienes lado a lado:","actions":[
  {"type":"show_chart","spec":{"kind":"bars","title":"BBVA","groupBy":"week","metric":"both","accountId":"<id BBVA>","dateFrom":"${today().slice(0,7)}-01","dateTo":"${today()}"}},
  {"type":"show_chart","spec":{"kind":"bars","title":"Santander","groupBy":"week","metric":"both","accountId":"<id Santander>","dateFrom":"${today().slice(0,7)}-01","dateTo":"${today()}"}}
]}

EJEMPLOS DE OTRAS ACCIONES:

Usuario: "ponle saldo inicial a Santander de 19,039"
Tu respuesta:
{"message":"Listo, dejé el saldo inicial de Santander en $19,039.","actions":[{"type":"edit_account","id":"<id real de Santander>","newInitialBalance":19039}]}

Usuario: "cambia el nombre de BBVA a Mi Banco"
Tu respuesta:
{"message":"Listo, BBVA ahora se llama Mi Banco.","actions":[{"type":"edit_account","id":"<id real de BBVA>","newName":"Mi Banco"}]}

Usuario: "el deposito de 7500 y el pago de seguro de 6800 son un reembolso, vincúlalos"
(asumiendo que en Movimientos recientes encuentras ingreso id="t-aa" de 7500 y gasto id="t-bb" de 6800)
Tu respuesta:
{"message":"Listo, los vinculé como reembolso. Solo cuenta como ingreso real la diferencia de $700.","actions":[{"type":"mark_passthrough","ids":["t-aa","t-bb"]}]}

Usuario: "el ingreso de los 7500 no es real, era un encargo"
(encuentras id="t-aa")
Tu respuesta:
{"message":"Marqué el ingreso de $7,500 como dinero de paso. No contará en tus estadísticas.","actions":[{"type":"mark_passthrough","ids":["t-aa"]}]}`;
}

/* persistencia — Firestore */
async function loadAll() {
  const u = auth.currentUser;
  if (!u) return { config: null, txs: [] };
  let config = null, txs = [];
  try {
    const snap = await getDoc(doc(db, "users", u.uid, "data", "config"));
    if (snap.exists()) config = snap.data().value;
  } catch (e) { console.error("loadAll config", e); }
  try {
    const snap = await getDoc(doc(db, "users", u.uid, "data", "txs"));
    if (snap.exists()) txs = snap.data().value || [];
  } catch (e) { console.error("loadAll txs", e); }
  return { config, txs };
}

async function persist(key, val) {
  const u = auth.currentUser;
  if (!u) return;
  const field = key === "cc:config" ? "config" : "txs";
  try {
    await setDoc(doc(db, "users", u.uid, "data", field), { value: val });
  } catch (e) { console.error("persist", e); }
}

/* llamada a Claude con imágenes (visión) */
async function callClaudeVision(system, userText, imagesB64) {
  const content = [
    ...imagesB64.map((b) => ({
      type: "image",
      source: { type: "base64", media_type: b.mediaType, data: b.data },
    })),
    { type: "text", text: userText },
  ];
  const body = { model: "claude-sonnet-4-6", max_tokens: 4000, system, messages: [{ role: "user", content }] };
  const isLocal = typeof window !== "undefined" && window.location.hostname === "localhost";
  let res;
  if (isLocal) {
    const key = import.meta.env.VITE_ANTHROPIC_KEY;
    res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01", "anthropic-dangerous-request-browser": "true" },
      body: JSON.stringify(body),
    });
  } else {
    res = await fetch("/api/claude", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  }
  if (!res.ok) throw new Error("vision " + res.status);
  const data = await res.json();
  return (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
}

/* leer File como base64 (sin el prefijo data:...) */
function fileToB64(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => {
      const result = String(r.result || "");
      const [meta, data] = result.split(",");
      const mediaType = (meta.match(/data:([^;]+)/) || [])[1] || "image/png";
      resolve({ mediaType, data });
    };
    r.onerror = () => reject(new Error("read failed"));
    r.readAsDataURL(file);
  });
}

/* llamada a Claude */
async function callClaude(system, messages) {
  const body = { model: "claude-sonnet-4-6", max_tokens: 1000, system, messages };
  const isLocal = typeof window !== "undefined" && window.location.hostname === "localhost";
  let res;
  if (isLocal) {
    const key = import.meta.env.VITE_ANTHROPIC_KEY;
    res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01", "anthropic-dangerous-request-browser": "true" },
      body: JSON.stringify(body),
    });
  } else {
    res = await fetch("/api/claude", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  }
  if (!res.ok) throw new Error("api " + res.status);
  const data = await res.json();
  return (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
}
function parseJSON(text) {
  const clean = text.replace(/```json/gi, "").replace(/```/g, "").trim();
  const a = clean.indexOf("{"), b = clean.lastIndexOf("}");
  return JSON.parse(a >= 0 && b >= 0 ? clean.slice(a, b + 1) : clean);
}

/* Categorizador compartido: keywords locales -> Claude.
 * Devuelve { catId, sure }. catId=null si no encuentra nada seguro. */
async function autoCategorize(desc, type, accountId, config) {
  const nd = norm(desc);
  const cats = config.categories.filter((c) => c.type === type && c.accountId === accountId);
  // 1) claves locales de las categorías de ESTA cuenta
  for (const c of cats) {
    for (const kw of c.keywords || []) {
      if (nd.includes(norm(kw))) return { catId: c.id, sure: true };
    }
  }
  // 2) claves aprendidas en otras cuentas con mismo nombre/tipo
  for (const c of cats) {
    const twins = config.categories.filter(
      (x) => x.id !== c.id && x.type === c.type && norm(x.name).trim() === norm(c.name).trim()
    );
    for (const twin of twins) {
      for (const kw of twin.keywords || []) {
        if (nd.includes(norm(kw))) return { catId: c.id, sure: true };
      }
    }
  }
  // 3) Claude
  try {
    const sys =
      'Eres un clasificador de transacciones personales. Responde SOLO con un objeto JSON, sin markdown ni texto extra: {"categoryId":"<id>" o null,"confident":true|false}. Devuelve null y confident:false si no estás razonablemente seguro.';
    const user = `Tipo: ${type === "income" ? "ingreso" : "gasto"}. Descripción: "${desc}". Categorías disponibles: ${JSON.stringify(
      cats.map((c) => ({ id: c.id, name: c.name }))
    )}`;
    const raw = await callClaude(sys, [{ role: "user", content: user }]);
    const p = parseJSON(raw);
    const valid = cats.some((c) => c.id === p.categoryId);
    if (valid && p.confident) return { catId: p.categoryId, sure: true };
    if (valid) return { catId: p.categoryId, sure: false };
  } catch (e) { /* sin API */ }
  return { catId: null, sure: false };
}

/* categorías por defecto para una cuenta nueva */
function defaultCatsForAccount(accId) {
  return DEFAULT_CATS.map((t) => ({
    id: uid(), name: t.name, emoji: t.emoji, type: t.type, accountId: accId,
    keywords: SEED_KW[norm(t.name).trim()] ? [...SEED_KW[norm(t.name).trim()]] : [],
  }));
}

/* construir config completa desde el chatbot o el set default */
function buildConfig(raw) {
  const accountMode = raw.accountMode === "multiple" ? "multiple" : "single";
  let accounts = (raw.accounts || []).map((a) => ({
    id: uid(), name: a.name || "Cuenta", initialBalance: Number(a.initialBalance) || 0,
  }));
  if (!accounts.length) accounts = [{ id: uid(), name: "Principal", initialBalance: 0 }];

  let templates = (raw.categories || []).map((c) => ({
    name: c.name || "Categoría", emoji: c.emoji || "📦",
    type: c.type === "income" ? "income" : "expense",
  }));
  if (!templates.some((c) => c.type === "income")) templates.push({ name: "Otros ingresos", emoji: "💰", type: "income" });
  if (!templates.some((c) => c.type === "expense")) templates.push({ name: "Otros gastos", emoji: "📦", type: "expense" });

  // cada cuenta arranca con su propia copia de las categorías
  const categories = [];
  accounts.forEach((acc) => {
    templates.forEach((t) => {
      categories.push({
        id: uid(), name: t.name, emoji: t.emoji, type: t.type, accountId: acc.id,
        keywords: SEED_KW[norm(t.name).trim()] ? [...SEED_KW[norm(t.name).trim()]] : [],
      });
    });
  });
  return { setupComplete: true, accountMode, accounts, categories };
}

/* migración: convierte categorías globales (versión vieja) en categorías por cuenta */
function migrate(config, txs) {
  if (!config || !config.categories) return { config, txs };
  let outConfig = config;
  let outTxs = txs || [];

  // Migración 1: categorías globales -> categorías por cuenta
  if (config.categories.some((c) => !c.accountId)) {
    const oldCats = config.categories;
    const newCats = [];
    const map = {};
    config.accounts.forEach((acc) => {
      oldCats.forEach((oc) => {
        const nid = uid();
        newCats.push({
          id: nid, name: oc.name, emoji: oc.emoji || "📦",
          type: oc.type === "income" ? "income" : "expense",
          accountId: acc.id, keywords: [...(oc.keywords || [])],
        });
        map[`${acc.id}|${oc.id}`] = nid;
      });
    });
    outTxs = outTxs.map((t) => {
      const nid = map[`${t.accountId}|${t.categoryId}`];
      return nid ? { ...t, categoryId: nid } : t;
    });
    outConfig = { ...config, categories: newCats };
  }

  // Migración 2: linkedTxId (uno a uno) -> groupId (muchos a muchos)
  const hasLinked = outTxs.some((t) => t.linkedTxId);
  if (hasLinked) {
    const byId = new Map(outTxs.map((t) => [t.id, t]));
    const groupOf = new Map(); // id -> groupId
    function find(id) {
      while (groupOf.get(id) && groupOf.get(id) !== id) id = groupOf.get(id);
      return id;
    }
    outTxs.forEach((t) => { if (t.passThrough) groupOf.set(t.id, t.id); });
    outTxs.forEach((t) => {
      if (!t.passThrough || !t.linkedTxId) return;
      const other = byId.get(t.linkedTxId);
      if (!other || !other.passThrough) return;
      const r1 = find(t.id), r2 = find(other.id);
      if (r1 !== r2) groupOf.set(r2, r1);
    });
    outTxs = outTxs.map((t) => {
      if (!t.passThrough) return t;
      const root = groupOf.get(t.id) ? find(t.id) : null;
      const { linkedTxId, ...rest } = t;
      return root ? { ...rest, groupId: t.groupId || root } : rest;
    });
  }

  // Migración 3: quitar pass-through del MVP — los movimientos viejos vuelven a contar como reales
  const hadPassThrough = outTxs.some((t) => t.passThrough || t.linkedTxId || t.groupId);
  if (hadPassThrough) {
    outTxs = outTxs.map((t) => {
      const { passThrough, linkedTxId, groupId, ...rest } = t;
      return rest;
    });
  }

  return { config: outConfig, txs: outTxs };
}

/* ========================================================================= */
/* =========================================================================
   AUTH — pantallas de bienvenida, login, registro, recuperar contraseña
   ========================================================================= */

const AUTH_INK = "#1B2230";
const AUTH_INK_SOFT = "#6B7585";
const AUTH_GREEN = "#1A7A6E";
const AUTH_LINE = "#DDE2E9";
const AUTH_CORAL = "#B5453A";

const COUNTRIES = [
  { name: "Mexico", flag: "🇲🇽" },
  { name: "United States", flag: "🇺🇸" },
  { name: "Argentina", flag: "🇦🇷" },
  { name: "Bolivia", flag: "🇧🇴" },
  { name: "Brazil", flag: "🇧🇷" },
  { name: "Canada", flag: "🇨🇦" },
  { name: "Chile", flag: "🇨🇱" },
  { name: "Colombia", flag: "🇨🇴" },
  { name: "Costa Rica", flag: "🇨🇷" },
  { name: "Cuba", flag: "🇨🇺" },
  { name: "Ecuador", flag: "🇪🇨" },
  { name: "El Salvador", flag: "🇸🇻" },
  { name: "España", flag: "🇪🇸" },
  { name: "Guatemala", flag: "🇬🇹" },
  { name: "Honduras", flag: "🇭🇳" },
  { name: "Nicaragua", flag: "🇳🇮" },
  { name: "Panama", flag: "🇵🇦" },
  { name: "Paraguay", flag: "🇵🇾" },
  { name: "Peru", flag: "🇵🇪" },
  { name: "Puerto Rico", flag: "🇵🇷" },
  { name: "Dominican Republic", flag: "🇩🇴" },
  { name: "Uruguay", flag: "🇺🇾" },
  { name: "Venezuela", flag: "🇻🇪" },
];

function ZafiLogo() {
  return (
    <div style={{ display:"flex", flexDirection:"column", alignItems:"center", gap:8, marginBottom:12 }}>
      <span style={{ fontFamily:"'Fraunces',serif", fontWeight:400, fontSize:38,
        letterSpacing:"-.05em", color:AUTH_INK, fontFeatureSettings:'"ss01"', lineHeight:1 }}>
        zafi
      </span>
      <div style={{ fontFamily:"'Montserrat',sans-serif", fontSize:11, color:AUTH_INK_SOFT, letterSpacing:".15em", fontWeight:500, textTransform:"uppercase" }}>
        {_lang === "es" ? "el futuro de tu dinero" : "the future of your money"}
      </div>
    </div>
  );
}

function AuthScreen() {
  const [screen, setScreen] = useState("welcome");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [name, setName] = useState("");
  const [gender, setGender] = useState("");
  const [age, setAge] = useState("");
  const [country, setCountry] = useState("Mexico");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [ok, setOk] = useState("");

  function reset() { setEmail(""); setPassword(""); setConfirm(""); setName(""); setGender(""); setAge(""); setCountry("Mexico"); setErr(""); setOk(""); }
  function go(s) { reset(); setScreen(s); }

  function ferr(code) {
    return ({
      "auth/email-already-in-use": "Ya existe una cuenta con ese correo.",
      "auth/invalid-email": "El correo no es válido.",
      "auth/weak-password": "La contraseña debe tener al menos 6 caracteres.",
      "auth/user-not-found": "No encontramos una cuenta con ese correo.",
      "auth/wrong-password": "Contraseña incorrecta.",
      "auth/invalid-credential": "Correo o contraseña incorrectos.",
      "auth/too-many-requests": "Demasiados intentos. Espera un momento.",
      "auth/network-request-failed": "Sin conexión a internet.",
    })[code] || "Algo salió mal. Intenta de nuevo.";
  }

  async function doRegister() {
    if (!name.trim()) { setErr(t("enterName")); return; }
    if (!gender) { setErr(t("selectGender")); return; }
    if (!age || Number(age) < 1 || Number(age) > 120) { setErr(t("enterValidAge")); return; }
    if (!country.trim()) { setErr(t("selectCountry")); return; }
    if (!email || !password || !confirm) { setErr(t("fillAllFields")); return; }
    if (password !== confirm) { setErr(t("passwordsDontMatch")); return; }
    if (password.length < 6) { setErr(t("passwordMinLength")); return; }
    setBusy(true); setErr("");
    try {
      const cred = await createUserWithEmailAndPassword(auth, email, password);
      // guardar perfil dentro del config para que no vuelva a pedir ProfileSetup
      const avatarId = defaultAvatarForGender(gender);
      try {
        await setDoc(doc(db, "users", cred.user.uid, "data", "config"), {
          value: {
            userName: name.trim(), userGender: gender, userAge: Number(age),
            userCountry: country.trim(), ...(avatarId ? { avatarId } : {}),
          }
        }, { merge: true });
      } catch (e) { /* no bloquea el registro */ }
    }
    catch(e) { setErr(ferr(e.code)); }
    finally { setBusy(false); }
  }

  async function doLogin() {
    if (!email || !password) { setErr("Llena todos los campos."); return; }
    setBusy(true); setErr("");
    try { await signInWithEmailAndPassword(auth, email, password); }
    catch(e) { setErr(ferr(e.code)); }
    finally { setBusy(false); }
  }

  async function doForgot() {
    if (!email) { setErr("Escribe tu correo primero."); return; }
    setBusy(true); setErr(""); setOk("");
    try { await sendPasswordResetEmail(auth, email); setOk("Te enviamos un correo para restablecer tu contraseña."); }
    catch(e) { setErr(ferr(e.code)); }
    finally { setBusy(false); }
  }

  const FONT = "'Montserrat', sans-serif";
  const inp = {
    width:"100%", padding:"14px 16px", borderRadius:16,
    border:"1px solid rgba(255,255,255,.5)", fontSize:15, fontFamily:FONT, fontWeight:400,
    background:"rgba(255,255,255,.55)", color:AUTH_INK, outline:"none",
    backdropFilter:"blur(10px)", WebkitBackdropFilter:"blur(10px)",
  };
  const btnP = {
    width:"100%", padding:15, borderRadius:99, border:"none",
    background:AUTH_INK, color:"#fff", fontSize:15, fontWeight:600,
    fontFamily:FONT, cursor:busy?"not-allowed":"pointer", opacity:busy?.6:1,
    letterSpacing:"-.01em", transition:"transform .1s ease",
  };
  const btnS = {
    width:"100%", padding:14, borderRadius:99,
    border:"1px solid rgba(255,255,255,.5)", background:"rgba(255,255,255,.45)",
    backdropFilter:"blur(10px)", WebkitBackdropFilter:"blur(10px)",
    color:AUTH_INK, fontSize:15, fontWeight:500,
    fontFamily:FONT, cursor:"pointer", letterSpacing:"-.01em",
  };
  const lnk = {
    background:"none", border:"none", color:"#5B6EE8",
    fontSize:14, fontFamily:FONT, cursor:"pointer",
    fontWeight:600, padding:0,
  };
  const lbl = {
    fontFamily:FONT, fontSize:12, fontWeight:600, color:AUTH_INK_SOFT,
    letterSpacing:".02em", marginBottom:6, display:"block",
  };
  const wrap = {
    minHeight:"100vh", display:"flex", flexDirection:"column",
    alignItems:"center", justifyContent:"center", padding:"40px 24px",
    background:"#DCE1E8", position:"relative", overflow:"hidden",
  };
  const blob = {
    position:"absolute", width:280, height:280, borderRadius:"50%",
    background:"radial-gradient(circle, rgba(96,165,250,.25) 0%, rgba(96,165,250,.05) 60%, transparent 80%)",
    filter:"blur(40px)", top:"15%", right:"-10%", pointerEvents:"none",
  };
  const blob2 = {
    position:"absolute", width:200, height:200, borderRadius:"50%",
    background:"radial-gradient(circle, rgba(167,139,250,.18) 0%, rgba(167,139,250,.04) 60%, transparent 80%)",
    filter:"blur(30px)", bottom:"20%", left:"-5%", pointerEvents:"none",
  };
  const box = {
    width:"100%", maxWidth:360, display:"flex", flexDirection:"column", gap:24,
  };

  if (screen === "welcome") return (
    <div style={{ minHeight:"100vh", position:"relative", overflow:"hidden", background:"#DCE1E8",
      display:"flex", flexDirection:"column", justifyContent:"flex-end" }}>
      {/* video de fondo */}
      <video
        ref={(el) => {
          if (el) {
            el.muted = true;
            const p = el.play();
            if (p && p.catch) p.catch(() => {});
          }
        }}
        autoPlay muted loop playsInline preload="auto"
        style={{ position:"absolute", inset:0, width:"100%", height:"100%",
          objectFit:"cover", zIndex:0, background:"#DCE1E8" }}
      >
        <source src="/zafi-intro.mp4" type="video/mp4" />
      </video>
      {/* degradado para legibilidad abajo */}
      <div style={{ position:"absolute", inset:0, zIndex:1, pointerEvents:"none",
        background:"linear-gradient(to bottom, rgba(220,225,232,0) 35%, rgba(220,225,232,.55) 70%, rgba(220,225,232,.85) 100%)" }} />

      {/* contenido abajo */}
      <div style={{ position:"relative", zIndex:2, padding:"0 26px calc(34px + env(safe-area-inset-bottom))",
        display:"flex", flexDirection:"column", gap:22 }}>
        <div>
          <div style={{ fontFamily:"'Fraunces',serif", fontWeight:400, fontSize:56,
            letterSpacing:"-.05em", color:AUTH_INK, lineHeight:1, fontFeatureSettings:'"ss01"' }}>zafi</div>
          <div style={{ fontFamily:"'Montserrat',sans-serif", fontWeight:400, fontSize:16,
            color:AUTH_INK_SOFT, marginTop:6, letterSpacing:"-.01em" }}>
            finanzas personales con <span style={{ fontFamily:"'Fraunces',serif", fontStyle:"italic" }}>IA</span>
          </div>
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          <button
            style={{ width:"100%", padding:16, borderRadius:99, border:"1px solid rgba(255,255,255,.6)",
              background:"rgba(255,255,255,.65)", backdropFilter:"blur(12px)", WebkitBackdropFilter:"blur(12px)",
              color:AUTH_INK, fontSize:15, fontWeight:600, fontFamily:"'Montserrat', sans-serif", cursor:"pointer",
              letterSpacing:"-.01em", boxShadow:"0 6px 24px rgba(30,40,60,.12)" }}
            onClick={() => go("method")}>{t("createAccount")}</button>
          <button
            style={{ width:"100%", padding:16, borderRadius:99, border:"1px solid rgba(255,255,255,.25)",
              background:"rgba(27,34,48,.55)", backdropFilter:"blur(12px)", WebkitBackdropFilter:"blur(12px)",
              color:"#fff", fontSize:15, fontWeight:500, fontFamily:"'Montserrat', sans-serif", cursor:"pointer",
              letterSpacing:"-.01em", boxShadow:"0 6px 24px rgba(30,40,60,.18)" }}
            onClick={() => go("login")}>{t("signIn")}</button>
        </div>
        <div style={{ fontSize:11.5, color:AUTH_INK_SOFT, textAlign:"center", lineHeight:1.5, opacity:.75, fontFamily:"'Montserrat', sans-serif" }}>
          Al crear una cuenta aceptas nuestros términos de uso y política de privacidad.
        </div>
      </div>
    </div>
  );

  // ===== Shared step page wrapper =====
  const stepWrap = {
    minHeight:"100vh", display:"flex", flexDirection:"column",
    background:"#EAEEF4", position:"relative", fontFamily:FONT,
  };
  const backBtn = {
    position:"absolute", top:20, left:20, zIndex:2, width:44, height:44,
    borderRadius:"50%", border:"none", background:"#fff",
    boxShadow:"0 2px 10px rgba(30,40,60,.1)", cursor:"pointer",
    fontSize:20, display:"flex", alignItems:"center", justifyContent:"center",
    color:AUTH_INK, fontFamily:FONT,
  };
  const stepBottom = {
    padding:"0 26px calc(28px + env(safe-area-inset-bottom))",
    display:"flex", flexDirection:"column", gap:12,
  };

  // ===== PASO 1: Método (Continue with email / Google / Apple) =====
  if (screen === "method") return (
    <div style={{ ...stepWrap, justifyContent:"flex-end" }}>
      <button style={backBtn} onClick={() => go("welcome")}>‹</button>
      <div style={{ flex:1 }} />
      <div style={{ padding:"0 26px 32px" }}>
        <div style={{ fontFamily:"'Fraunces',serif", fontWeight:400, fontSize:64,
          letterSpacing:"-.05em", color:AUTH_INK, lineHeight:1, fontFeatureSettings:'"ss01"' }}>zafi</div>
      </div>
      <div style={stepBottom}>
        <button style={{ ...btnP, borderRadius:99 }} onClick={() => { setErr(""); setScreen("country"); }}>
          {t("continueWithEmail")}
        </button>
        <button style={{ width:"100%", padding:15, borderRadius:99, border:"none",
          background:"#4285F4", color:"#fff", fontSize:15, fontWeight:600,
          fontFamily:FONT, cursor:"pointer", letterSpacing:"-.01em",
          display:"flex", alignItems:"center", justifyContent:"center", gap:10 }}
          onClick={async () => {
            setBusy(true); setErr("");
            try {
              const gp = new GoogleAuthProvider();
              await signInWithPopup(auth, gp);
            }
            catch(e) { setErr(ferr(e.code)); }
            setBusy(false);
          }}
          disabled={busy}>
          {t("continueWithGoogle")}
          <span style={{ fontSize:18, fontWeight:700 }}>G</span>
        </button>
        <button style={{ width:"100%", padding:15, borderRadius:99, border:"none",
          background:"#000", color:"#fff", fontSize:15, fontWeight:600,
          fontFamily:FONT, cursor:"pointer", letterSpacing:"-.01em",
          display:"flex", alignItems:"center", justifyContent:"center", gap:10 }}
          onClick={async () => {
            setBusy(true); setErr("");
            try {
              const ap = new OAuthProvider("apple.com");
              await signInWithPopup(auth, ap);
            }
            catch(e) { setErr(ferr(e.code)); }
            setBusy(false);
          }}
          disabled={busy}>
          {t("continueWithApple")}
          <span style={{ fontSize:18 }}></span>
        </button>
        {err && <div style={{ fontSize:13, color:AUTH_CORAL, fontWeight:500, textAlign:"center" }}>{err}</div>}
      </div>
    </div>
  );

  // ===== PASO 2: País =====
  const selCountry = COUNTRIES.find(c => c.name === country) || COUNTRIES[0];
  if (screen === "country") return (
    <div style={{ ...stepWrap }}>
      <button style={backBtn} onClick={() => { setErr(""); setScreen("method"); }}>‹</button>
      <div style={{ flex:1, padding:"90px 26px 0" }}>
        <div style={{ fontSize:28, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", lineHeight:1.2 }}>
          {t("countryOfResidence")}
        </div>
        <div style={{ fontSize:14, color:AUTH_INK_SOFT, marginTop:8, lineHeight:1.6 }}>
          {t("countryDesc")}
        </div>
        <div style={{ marginTop:32 }}>
          <div style={{ fontSize:13, color:AUTH_INK_SOFT, marginBottom:8 }}>{t("country")}</div>
          <div style={{ position:"relative" }}>
            <select
              value={country}
              onChange={e => setCountry(e.target.value)}
              style={{ width:"100%", padding:"12px 0", fontSize:16, fontWeight:600,
                fontFamily:FONT, color:"transparent", background:"transparent",
                border:"none", borderBottom:"1px solid rgba(27,34,48,.15)",
                outline:"none", appearance:"none", cursor:"pointer",
                position:"relative", zIndex:1 }}>
              {COUNTRIES.map(c => (
                <option key={c.name} value={c.name}>{c.flag}  {c.name}</option>
              ))}
            </select>
            <div style={{ position:"absolute", top:12, left:0, fontSize:16, fontWeight:600, color:AUTH_INK, pointerEvents:"none" }}>
              {selCountry.flag} {selCountry.name}
            </div>
          </div>
        </div>
      </div>
      <div style={stepBottom}>
        <div style={{ fontSize:13, color:AUTH_INK_SOFT, lineHeight:1.6, marginBottom:4 }}>
          {t("legalConsent")} <span style={{ textDecoration:"underline", cursor:"pointer" }}>{t("termsAndConditions")}</span> {_lang==="es"?"y":"and"} <span style={{ textDecoration:"underline", cursor:"pointer" }}>{t("privacyPolicy")}</span>, {t("legalConsentEnd")}
        </div>
        <button style={{ ...btnP, borderRadius:99 }} onClick={() => { setErr(""); setScreen("profile"); }}>
          {t("getStarted")}
        </button>
      </div>
    </div>
  );

  // ===== PASO 3: Datos personales (Nombre, Edad, Género) =====
  if (screen === "profile") return (
    <div style={{ ...stepWrap }}>
      <button style={backBtn} onClick={() => { setErr(""); setScreen("country"); }}>‹</button>
      <div style={{ flex:1, padding:"90px 26px 0" }}>
        <div style={{ fontSize:28, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", lineHeight:1.2 }}>
          {t("aboutYou")}
        </div>
        <div style={{ fontSize:14, color:AUTH_INK_SOFT, marginTop:8, lineHeight:1.6 }}>
          {t("tellUsAboutYou")}
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:18, marginTop:28 }}>
          <div>
            <label style={lbl}>{t("name")}</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)" }} type="text" placeholder={t("yourName")} value={name} onChange={e=>setName(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>{t("age")}</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)", width:120 }} type="text" inputMode="numeric" placeholder="00" value={age}
              onChange={e=>setAge(e.target.value.replace(/[^0-9]/g,"").slice(0,3))} />
          </div>
          <div>
            <label style={lbl}>{t("gender")}</label>
            <div style={{ display:"flex", gap:9 }}>
              {[["male",t("male")],["female",t("female")],["other",t("other")]].map(([k,l])=>(
                <button key={k} type="button" onClick={()=>setGender(k)}
                  style={{ flex:1, padding:"13px 8px", borderRadius:14, cursor:"pointer", fontFamily:FONT,
                    fontSize:14, fontWeight:gender===k?600:400,
                    background: gender===k ? AUTH_INK : "rgba(255,255,255,.7)",
                    color: gender===k ? "#fff" : AUTH_INK,
                    border:`1px solid ${gender===k ? AUTH_INK : "rgba(0,0,0,.08)"}`,
                    transition:"background .15s, color .15s" }}>
                  {l}
                </button>
              ))}
            </div>
          </div>
        </div>
        {err && <div style={{ fontSize:13, color:AUTH_CORAL, fontWeight:500, marginTop:12 }}>{err}</div>}
      </div>
      <div style={stepBottom}>
        <button style={{ ...btnP, borderRadius:99 }}
          onClick={() => {
            if (!name.trim()) { setErr(t("enterName")); return; }
            if (!age || Number(age) < 1 || Number(age) > 120) { setErr(t("enterValidAge")); return; }
            if (!gender) { setErr(t("selectGender")); return; }
            setErr(""); setScreen("credentials");
          }}>
          Continue
        </button>
      </div>
    </div>
  );

  // ===== PASO 4: Credenciales (Email + Contraseña) =====
  if (screen === "credentials") return (
    <div style={{ ...stepWrap }}>
      <button style={backBtn} onClick={() => { setErr(""); setScreen("profile"); }}>‹</button>
      <div style={{ flex:1, padding:"90px 26px 0" }}>
        <div style={{ fontSize:28, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", lineHeight:1.2 }}>
          {t("createAccount")}
        </div>
        <div style={{ fontSize:14, color:AUTH_INK_SOFT, marginTop:8, lineHeight:1.6 }}>
          {t("credentialsDesc")}
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:16, marginTop:28 }}>
          <div>
            <label style={lbl}>{t("email")}</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)" }} type="email" placeholder="you@example.com" value={email} onChange={e=>setEmail(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>{t("password")}</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)" }} type="password" placeholder={_lang==="es"?"Mínimo 6 caracteres":"At least 6 characters"} value={password} onChange={e=>setPassword(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>{t("confirmPassword")}</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)" }} type="password" placeholder={_lang==="es"?"Repite tu contraseña":"Repeat your password"} value={confirm} onChange={e=>setConfirm(e.target.value)} />
          </div>
        </div>
        {err && <div style={{ fontSize:13, color:AUTH_CORAL, fontWeight:500, marginTop:12 }}>{err}</div>}
      </div>
      <div style={stepBottom}>
        <button style={{ ...btnP, borderRadius:99 }} onClick={doRegister} disabled={busy}>
          {busy ? (_lang==="es" ? "Creando cuenta…" : "Creating account…") : t("createAccount")}
        </button>
      </div>
    </div>
  );

  if (screen === "login") return (
    <div style={wrap}>
      <div style={blob} />
      <div style={blob2} />
      <div style={{ ...box, position:"relative", zIndex:1 }}>
        <div style={{ marginBottom: 24 }}>
          <ZafiLogo />
        </div>
        <div>
          <div style={{ fontFamily:FONT, fontSize:24, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", marginBottom:4 }}>{t("welcomeBack")}</div>
          <div style={{ fontFamily:FONT, fontSize:14, fontWeight:400, color:AUTH_INK_SOFT }}>{t("signInToContinue")}</div>
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
          <div>
            <label style={lbl}>{t("email")}</label>
            <input style={inp} type="email" placeholder={_lang === "es" ? "tucorreo@ejemplo.com" : "you@example.com"} value={email} onChange={e=>setEmail(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>{t("password")}</label>
            <input style={inp} type="password" placeholder={_lang === "es" ? "Tu contraseña" : "Your password"} value={password} onChange={e=>setPassword(e.target.value)} />
          </div>
        </div>
        {err && <div style={{ fontFamily:FONT, fontSize:13.5, color:AUTH_CORAL, fontWeight:500 }}>{err}</div>}
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          <button style={btnP} onClick={doLogin} disabled={busy}>{busy ? (_lang === "es" ? "Entrando…" : "Signing in…") : t("signIn")}</button>
          <button style={btnS} onClick={() => go("welcome")}>{_lang === "es" ? "← Regresar" : "← Back"}</button>
        </div>
        <div style={{ textAlign:"center", fontSize:14, color:AUTH_INK_SOFT, display:"flex", flexDirection:"column", gap:8 }}>
          <button style={lnk} onClick={() => go("forgot")}>{t("forgotPassword")}</button>
          <span style={{ fontFamily:FONT }}>{_lang === "es" ? "¿No tienes cuenta?" : "Don't have an account?"}{" "}<button style={lnk} onClick={() => go("method")}>{t("createAccount")}</button></span>
        </div>
      </div>
    </div>
  );

  if (screen === "forgot") return (
    <div style={wrap}>
      <div style={blob} />
      <div style={blob2} />
      <div style={{ ...box, position:"relative", zIndex:1 }}>
        <div style={{ marginBottom: 24 }}>
          <ZafiLogo />
        </div>
        <div>
          <div style={{ fontFamily:FONT, fontSize:24, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", marginBottom:4 }}>{t("resetPassword")}</div>
          <div style={{ fontFamily:FONT, fontSize:14, fontWeight:400, color:AUTH_INK_SOFT, lineHeight:1.5 }}>{t("resetPasswordDesc")}</div>
        </div>
        <div>
          <label style={lbl}>{t("email")}</label>
          <input style={inp} type="email" placeholder={_lang === "es" ? "tucorreo@ejemplo.com" : "you@example.com"} value={email} onChange={e=>setEmail(e.target.value)} />
        </div>
        {err && <div style={{ fontFamily:FONT, fontSize:13.5, color:AUTH_CORAL, fontWeight:500 }}>{err}</div>}
        {ok && <div style={{ fontFamily:FONT, fontSize:13.5, color:"#5B6EE8", fontWeight:500 }}>{ok}</div>}
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          <button style={btnP} onClick={doForgot} disabled={busy}>{busy ? (_lang === "es" ? "Enviando…" : "Sending…") : t("sendLink")}</button>
          <button style={btnS} onClick={() => go("login")}>{_lang === "es" ? "← Regresar" : "← Back"}</button>
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   APP — componente principal
   ========================================================================= */
/* ===================== PROFILE SETUP (post-auth for Google/Apple) ======= */
function ProfileSetup({ user, config, saveConfig, onDone }) {
  const FONT = "'Montserrat', sans-serif";
  const [name, setName] = useState("");
  const [age, setAge] = useState("");
  const [gender, setGender] = useState("");
  const [country, setCountry] = useState("Mexico");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const selCountry = COUNTRIES.find(c => c.name === country) || COUNTRIES[0];

  const save = async () => {
    if (!name.trim()) { setErr(t("enterName")); return; }
    if (!age || Number(age) < 1) { setErr(t("enterValidAge")); return; }
    if (!gender) { setErr(t("selectGender")); return; }
    setBusy(true); setErr("");
    try {
      const avatarId = defaultAvatarForGender(gender);
      const profileData = {
        userName: name.trim(), userGender: gender, userAge: Number(age),
        userCountry: country, ...(avatarId ? { avatarId } : {}),
      };
      // merge into existing config or create minimal config
      const updated = { ...(config || {}), ...profileData };
      saveConfig(updated);
      onDone();
    } catch (e) { setErr(t("somethingWentWrong")); }
    setBusy(false);
  };

  const stepWrap = {
    minHeight:"100vh", display:"flex", flexDirection:"column",
    background:"#EAEEF4", fontFamily:FONT,
  };
  const lbl = { fontFamily:FONT, fontSize:12, fontWeight:600, color:"#6B7585", letterSpacing:".02em", marginBottom:6, display:"block" };
  const inp = { width:"100%", padding:"14px 16px", borderRadius:16,
    border:"1px solid rgba(0,0,0,.06)", fontSize:15, fontFamily:FONT, fontWeight:400,
    background:"rgba(255,255,255,.7)", color:"#1B2230", outline:"none" };

  return (
    <div style={stepWrap}>
      <div style={{ flex:1, padding:"100px 26px 0" }}>
        <div style={{ fontSize:28, fontWeight:700, color:"#1B2230", letterSpacing:"-.02em", lineHeight:1.2 }}>
          Welcome to Zafi
        </div>
        <div style={{ fontSize:14, color:"#6B7585", marginTop:8, lineHeight:1.6 }}>
          {t("tellUsAboutYou")}
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:18, marginTop:28 }}>
          <div>
            <label style={lbl}>Your name</label>
            <input style={inp} type="text" placeholder={t("howToCallYou")} value={name} onChange={e=>setName(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>{t("age")}</label>
            <input style={{ ...inp, width:120 }} type="text" inputMode="numeric" placeholder="00" value={age}
              onChange={e=>setAge(e.target.value.replace(/[^0-9]/g,"").slice(0,3))} />
          </div>
          <div>
            <label style={lbl}>{t("gender")}</label>
            <div style={{ display:"flex", gap:9 }}>
              {[["male",t("male")],["female",t("female")],["other",t("other")]].map(([k,l])=>(
                <button key={k} type="button" onClick={()=>setGender(k)}
                  style={{ flex:1, padding:"13px 8px", borderRadius:14, cursor:"pointer", fontFamily:FONT,
                    fontSize:14, fontWeight:gender===k?600:400,
                    background: gender===k ? "#1B2230" : "rgba(255,255,255,.7)",
                    color: gender===k ? "#fff" : "#1B2230",
                    border:`1px solid ${gender===k ? "#1B2230" : "rgba(0,0,0,.08)"}`,
                    transition:"background .15s, color .15s" }}>
                  {l}
                </button>
              ))}
            </div>
          </div>
          <div>
            <label style={lbl}>{t("country")}</label>
            <div style={{ position:"relative" }}>
              <select value={country} onChange={e => setCountry(e.target.value)}
                style={{ width:"100%", padding:"12px 0", fontSize:16, fontWeight:600,
                  fontFamily:FONT, color:"transparent", background:"transparent",
                  border:"none", borderBottom:"1px solid rgba(27,34,48,.15)",
                  outline:"none", appearance:"none", cursor:"pointer", position:"relative", zIndex:1 }}>
                {COUNTRIES.map(c => <option key={c.name} value={c.name}>{c.flag}  {c.name}</option>)}
              </select>
              <div style={{ position:"absolute", top:12, left:0, fontSize:16, fontWeight:600, color:"#1B2230", pointerEvents:"none" }}>
                {selCountry.flag} {selCountry.name}
              </div>
            </div>
          </div>
        </div>
        {err && <div style={{ fontSize:13, color:"#B5453A", fontWeight:500, marginTop:12 }}>{err}</div>}
      </div>
      <div style={{ padding:"0 26px calc(28px + env(safe-area-inset-bottom))" }}>
        <button style={{ width:"100%", padding:15, borderRadius:99, border:"none",
          background:"#1B2230", color:"#fff", fontSize:15, fontWeight:600,
          fontFamily:FONT, cursor:busy?"not-allowed":"pointer", opacity:busy?.6:1 }}
          onClick={save} disabled={busy}>
          {busy ? t("saving") : t("continueBtn")}
        </button>
      </div>
    </div>
  );
}

/* ===================== SPLASH SCREEN ==================================== */
function SplashScreen({ onDone }) {
  useEffect(() => {
    const t = setTimeout(onDone, 1800);
    return () => clearTimeout(t);
  }, [onDone]);

  return (
    <div className="cc-splash">
      <style>{STYLE}</style>
      <div className="cc-splash-word">zafi</div>
      <div className="cc-splash-tag">{_lang === "es" ? "el futuro de tu dinero" : "the future of your money"}</div>
    </div>
  );
}

export default function App() {
  const [splashDone, setSplashDone] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const [config, setConfig] = useState(null);
  const [txs, setTxs] = useState([]);
  const [toast, setToast] = useState(null);
  const [user, setUser] = useState(undefined); // undefined=cargando, null=no logueado
  const [profileDone, setProfileDone] = useState(false); // did user complete profile?

  useEffect(() => {
    if (typeof document !== "undefined") document.title = "Zafi · Finanzas personales con IA";
  }, []);

  // Escuchar sesión de Firebase
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => setUser(u));
    return () => unsub();
  }, []);

  // Cargar datos locales — siempre se define, solo corre cuando hay usuario
  useEffect(() => {
    if (!user) { setProfileDone(false); setLoaded(false); return; }
    let cancelled = false;
    (async () => {
      try {
        const { config: c, txs: t } = await loadAll();
        if (cancelled) return;
        if (c) {
          const m = migrate(c, t || []);
          const r = processRecurring(m.config, m.txs);
          setConfig(r.config);
          setTxs(r.txs);
          const configChanged = m.config !== c || r.generated > 0;
          const txsChanged = m.txs !== t || r.generated > 0;
          if (configChanged) persist("cc:config", r.config);
          if (txsChanged) persist("cc:txs", r.txs);
          if (r.generated > 0) {
            setTimeout(() => {
              setToast(`Se aplicaron ${r.generated} movimiento${r.generated === 1 ? "" : "s"} de recurrencias`);
              setTimeout(() => setToast(null), 2800);
            }, 600);
          }
        } else if (t) {
          setTxs(t);
        }
        // Check if user has completed profile (stored in config)
        // Also accept existing users who have setupComplete but no userName yet
        const hasProfile = !!(c && (c.userName || c.setupComplete));
        if (!cancelled) setProfileDone(hasProfile);
      } catch (e) {
        console.error("load error", e);
        if (!cancelled) setProfileDone(false);
      }
      if (!cancelled) setLoaded(true);
    })();
    return () => { cancelled = true; };
  }, [user]); // re-corre cuando cambia el usuario

  // Global: deslizar hacia abajo para cerrar cualquier modal (cc-sheet)
  useEffect(() => {
    let startY = 0, currentDy = 0, sheet = null, dragging = false;
    const onStart = (e) => {
      const grip = e.target.closest(".cc-grip");
      if (!grip) return;
      sheet = grip.closest(".cc-sheet");
      if (!sheet) return;
      startY = e.touches[0].clientY;
      currentDy = 0;
      dragging = true;
      sheet.style.transition = "none";
    };
    const onMove = (e) => {
      if (!dragging || !sheet) return;
      const dy = e.touches[0].clientY - startY;
      if (dy < 0) { currentDy = 0; return; }
      currentDy = dy;
      sheet.style.transform = `translateY(${dy}px)`;
      sheet.style.opacity = Math.max(0.3, 1 - dy / 400);
    };
    const onEnd = () => {
      if (!dragging || !sheet) return;
      dragging = false;
      sheet.style.transition = "transform .25s ease, opacity .2s ease";
      if (currentDy > 100) {
        sheet.style.transform = "translateY(100vh)";
        sheet.style.opacity = "0";
        // click en el overlay para cerrar
        setTimeout(() => {
          const overlay = sheet.closest(".cc-overlay");
          if (overlay) overlay.click();
          // reset
          sheet.style.transform = "";
          sheet.style.opacity = "";
          sheet.style.transition = "";
        }, 220);
      } else {
        sheet.style.transform = "";
        sheet.style.opacity = "";
      }
      sheet = null;
    };
    document.addEventListener("touchstart", onStart, { passive: true });
    document.addEventListener("touchmove", onMove, { passive: true });
    document.addEventListener("touchend", onEnd);
    return () => {
      document.removeEventListener("touchstart", onStart);
      document.removeEventListener("touchmove", onMove);
      document.removeEventListener("touchend", onEnd);
    };
  }, []);

  // Detectar tema del sistema para modo "auto"
  const [systemDark, setSystemDark] = useState(() =>
    typeof window !== "undefined" && window.matchMedia?.("(prefers-color-scheme: dark)").matches);
  useEffect(() => {
    const mq = window.matchMedia?.("(prefers-color-scheme: dark)");
    if (!mq) return;
    const handler = (e) => setSystemDark(e.matches);
    mq.addEventListener("change", handler);
    return () => mq.removeEventListener("change", handler);
  }, []);

  const themeMode = config?.theme || "auto";
  const isDarkTheme = themeMode === "dark" || (themeMode === "auto" && systemDark);
  const bgVideoSrc = isDarkTheme ? "/zafi-bg-dark.mp4" : "/zafi-bg.mp4";

  // Body background para tema oscuro/claro
  useEffect(() => {
    const bg = isDarkTheme ? "#0D0F14" : "#DCE1E8";
    document.body.style.backgroundColor = bg;
    document.documentElement.style.backgroundColor = bg;
  }, [isDarkTheme]);

  // Splash de bienvenida — siempre primero al abrir la app
  if (!splashDone) return <SplashScreen onDone={() => setSplashDone(true)} />;


  // Pantalla de carga mientras Firebase verifica sesión
  if (user === undefined) return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center",
      justifyContent:"center", background:"#DCE1E8" }}>
      <div style={{ display:"flex", flexDirection:"column", alignItems:"center", gap:12 }}>
        <span style={{ fontFamily:"'Fraunces',serif", fontWeight:400, fontSize:26,
          letterSpacing:"-.05em", color:"#1B2230" }}>zafi</span>
        <div style={{ fontSize:13, color:"#A4ACBA" }}>Cargando…</div>
      </div>
    </div>
  );

  // Sin sesión → pantalla de Auth
  if (!user) return <AuthScreen />;

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2400);
  };
  const saveConfig = (c) => { setConfig(c); persist("cc:config", c); };
  const saveTxs = (t) => { setTxs(t); persist("cc:txs", t); };
  const resetAll = async () => {
    const u = auth.currentUser;
    if (u) {
      try { await deleteDoc(doc(db, "users", u.uid, "data", "config")); } catch (e) {}
      try { await deleteDoc(doc(db, "users", u.uid, "data", "txs")); } catch (e) {}
    }
    setConfig(null);
    setTxs([]);
    showToast("Listo, empezamos desde cero");
  };

  if (!loaded)
    return (
      <div className={`cc-root ${isDarkTheme ? "cc-dark" : ""}`} style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh" }}>
        <style>{STYLE}</style>
        <div className="cc-video-bg">
          <video src={bgVideoSrc} autoPlay muted loop playsInline preload="auto" key={bgVideoSrc} />
        </div>
        <div className="cc-dots"><span /><span /><span /></div>
      </div>
    );

  // Si el usuario no tiene perfil (Google/Apple sin nombre), pedir datos
  if (!profileDone)
    return <ProfileSetup user={user} config={config} saveConfig={saveConfig} onDone={() => setProfileDone(true)} />;

  return (
    <div className={`cc-root ${isDarkTheme ? "cc-dark" : ""}`}>
      <style>{STYLE}</style>
      <div className="cc-video-bg">
        <video src={bgVideoSrc} autoPlay muted loop playsInline preload="auto" key={bgVideoSrc} />
      </div>
      <div className="cc-bg-wave" />
      {!config?.setupComplete ? (
        <Onboarding onDone={(built) => saveConfig({ ...config, ...built })} />
      ) : (
        <Main config={config} txs={txs} saveConfig={saveConfig} saveTxs={saveTxs} showToast={showToast} resetAll={resetAll} />
      )}
      {toast && <div className="cc-toast">{toast}</div>}
    </div>
  );
}

/* ============================= ONBOARDING ================================ */
function Onboarding({ onDone }) {
  const GREETING =
    "¡Hola! Bienvenido a Zafi 👋\n\nSoy tu asistente. Te ayudaré a configurar la app en menos de un minuto, platicando aquí mismo.\n\n¿Quieres que use un set de categorías recomendado, o prefieres armarlas tú?";
  const [msgs, setMsgs] = useState([{ role: "bot", text: GREETING }]);
  const [suggs, setSuggs] = useState(["Usa las recomendadas", "Quiero armarlas yo"]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const [apiOk, setApiOk] = useState(true);
  const history = useRef([]); // [{role, content}]
  const scroller = useRef(null);

  useEffect(() => {
    if (scroller.current) scroller.current.scrollTop = scroller.current.scrollHeight;
  }, [msgs, busy]);

  const SYSTEM = `Eres el asistente de configuración de "Zafi", una app de finanzas personales con IA. El usuario habla español de México. Conversa de forma cálida, breve y natural.

Debes averiguar 2 cosas:
1) CATEGORÍAS para clasificar ingresos y gastos. Set recomendado: Sueldo, Negocio, Freelance, Otros ingresos (ingresos); Súper/Despensa, Restaurantes, Transporte/Gasolina, Casa/Renta, Servicios, Salud, Entretenimiento, Suscripciones, Otros gastos (gastos).
2) CUENTAS: si maneja una sola cuenta para todo, o varias (efectivo, banco, tarjeta, etc.).

REGLAS DE RESPUESTA:
- Responde SIEMPRE con UN SOLO objeto JSON válido, sin markdown ni texto fuera del JSON.
- Formato: {"message":"texto para el usuario","suggestions":["opción corta",...],"done":false,"config":null}
- "message": 2-4 frases máximo, UNA pregunta a la vez.
- "suggestions": hasta 4 respuestas rápidas cortas que el usuario puede tocar. Omite o usa [] si no aplica.
- Cuando ya tengas categorías y modo de cuentas definidos, pon "done":true y llena "config":
  {"accountMode":"single"|"multiple","accounts":[{"name":"...","initialBalance":<número: dinero que ya tiene en esa cuenta>}],"categories":[{"name":"...","emoji":"emoji","type":"income"|"expense"}]}
- En modo "single" usa una cuenta llamada "Principal" salvo que el usuario diga otro nombre.
- Si el usuario maneja varias cuentas, pregúntale el saldo inicial (dinero que tiene ahorita) de cada una y ponlo en "initialBalance". Si no lo sabe, usa 0.
- Cada categoría necesita un emoji apropiado.
- Si el usuario pide usar lo recomendado, ve directo a done:true con el set recomendado y pregunta el modo de cuentas si aún no lo sabes.`;

  async function send(text) {
    const userText = (text ?? input).trim();
    if (!userText || busy) return;
    setMsgs((m) => [...m, { role: "me", text: userText }]);
    setInput("");
    setSuggs([]);
    setBusy(true);
    history.current.push({ role: "user", content: userText });
    try {
      const raw = await callClaude(SYSTEM, history.current);
      const parsed = parseJSON(raw);
      history.current.push({ role: "assistant", content: parsed.message || raw });
      setMsgs((m) => [...m, { role: "bot", text: parsed.message || "Listo." }]);
      setSuggs(Array.isArray(parsed.suggestions) ? parsed.suggestions.slice(0, 4) : []);
      if (parsed.done && parsed.config) {
        setTimeout(() => onDone(buildConfig(parsed.config)), 700);
      }
    } catch (e) {
      setApiOk(false);
      setMsgs((m) => [
        ...m,
        { role: "bot", text: "Tuve un problema para conectar con la IA. No hay bronca: puedes usar la configuración recomendada con el botón de abajo. 👇" },
      ]);
    }
    setBusy(false);
  }

  return (
    <div>
      <div className="cc-top">
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 2 }}>
          <div className="cc-logo"><span className="cc-logo-dot" />zafi</div>
          <div style={{ fontSize: 11, color: "var(--ink-soft)", marginLeft: 19, letterSpacing: "-.005em" }}>
            Finanzas personales con IA
          </div>
        </div>
      </div>
      <div className="cc-wrap" style={{ paddingBottom: 40 }}>
        <div
          ref={scroller}
          style={{ display: "flex", flexDirection: "column", gap: 10, padding: "8px 2px 14px", maxHeight: "58vh", overflowY: "auto" }}
        >
          {msgs.map((m, i) => (
            <div key={i} className={`cc-bubble ${m.role} cc-fade`} style={{ whiteSpace: "pre-wrap" }}>
              {m.text}
            </div>
          ))}
          {busy && (
            <div className="cc-bubble bot"><span className="cc-dots"><span /><span /><span /></span></div>
          )}
        </div>

        {suggs.length > 0 && !busy && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
            {suggs.map((s, i) => (
              <button key={i} className="cc-chip" onClick={() => send(s)}>{s}</button>
            ))}
          </div>
        )}

        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <input
            className="cc-input"
            placeholder="Escribe tu respuesta…"
            value={input}
            disabled={busy}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && send()}
          />
          <button className="cc-btn cc-btn-primary" disabled={busy || !input.trim()} onClick={() => send()}>
            Enviar
          </button>
        </div>

        <div style={{ textAlign: "center" }}>
          <button
            className="cc-btn"
            style={{ fontSize: 13 }}
            onClick={() => onDone(buildConfig({ accountMode: "single", accounts: [{ name: "Principal" }], categories: DEFAULT_CATS }))}
          >
            {apiOk ? "Saltar — usar configuración recomendada" : "Usar configuración recomendada ahora"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* =============================== MAIN ==================================== */
function Main({ config, txs, saveConfig, saveTxs, showToast, resetAll }) {
  setAppLang(config.language || "es");

  // Parallax del fondo de video (con límite para no mostrar el borde)
  useEffect(() => {
    const onScroll = () => {
      // Video es 180% del viewport, escalado 1.06 = ~190%
      // Extra space: 90% del vh, dividido 2 = 45% arriba y abajo
      // Max translateY antes de mostrar el borde: ~40% del vh
      const maxY = window.innerHeight * 0.38;
      const raw = window.scrollY * 0.15; // velocidad más sutil
      const clamped = Math.min(raw, maxY);
      document.documentElement.style.setProperty("--parallax-y", `${-clamped}px`);
    };
    window.addEventListener("scroll", onScroll, { passive: true });
    return () => window.removeEventListener("scroll", onScroll);
  }, []);
  const [tab, setTab] = useState("inicio");
  const [adding, setAdding] = useState(false);
  const [editingTx, setEditingTx] = useState(null);
  const [chatOpen, setChatOpen] = useState(false);
  const [chatVoice, setChatVoice] = useState(false);
  const [importOpen, setImportOpen] = useState(false);
  const [addMenuOpen, setAddMenuOpen] = useState(false);
  const [accountsOpen, setAccountsOpen] = useState(false);
  const [excelOpen, setExcelOpen] = useState(false);
  const [rangeOpen, setRangeOpen] = useState(false);
  const [recurringOpen, setRecurringOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);

  const dateRange = config.dateRange || DEFAULT_RANGE;
  const setDateRange = (newRange) => {
    saveConfig({ ...config, dateRange: newRange });
  };

  const balance = grandTotal(config, txs);

  // === Motor de recurrentes: al cargar, genera movimientos pendientes ===
  const recurringRanRef = useRef(false);
  useEffect(() => {
    if (recurringRanRef.current) return;
    const rules = config.recurring || [];
    if (!rules.length) return;
    const { newTxs, updatedRecurring } = runRecurringRules(rules);
    if (newTxs.length > 0) {
      recurringRanRef.current = true;
      saveTxs([...newTxs, ...txs]);
      saveConfig({ ...config, recurring: updatedRecurring });
      showToast(`${newTxs.length} movimiento${newTxs.length === 1 ? "" : "s"} recurrente${newTxs.length === 1 ? "" : "s"} generado${newTxs.length === 1 ? "" : "s"}`);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [config.recurring]);

  const saveRecurring = (rules) => {
    // al guardar reglas, corre inmediatamente las que ya tengan fechas vencidas
    const { newTxs, updatedRecurring } = runRecurringRules(rules);
    if (newTxs.length > 0) {
      saveTxs([...newTxs, ...txs]);
      showToast(`${newTxs.length} movimiento${newTxs.length === 1 ? "" : "s"} recurrente${newTxs.length === 1 ? "" : "s"} generado${newTxs.length === 1 ? "" : "s"}`);
    }
    saveConfig({ ...config, recurring: updatedRecurring });
  };

  const upsertTx = (tx, learnedCats, linkInfo) => {
    const exists = txs.some((t) => t.id === tx.id);
    let nextTxs = exists ? txs.map((t) => (t.id === tx.id ? tx : t)) : [tx, ...txs];

    // Si nos pasaron información de vinculación, también marcamos a los demás
    // miembros del grupo como passThrough con el mismo groupId
    if (linkInfo && linkInfo.linkIds && linkInfo.linkIds.length && linkInfo.groupIdToUse) {
      const ids = new Set(linkInfo.linkIds);
      nextTxs = nextTxs.map((t) =>
        ids.has(t.id)
          ? { ...t, passThrough: true, groupId: linkInfo.groupIdToUse, categoryId: null }
          : t
      );
    }

    // Si dejamos de ser passThrough y antes estábamos en un grupo:
    // quitamos también del grupo a los huérfanos del grupo si quedan menos de 2
    if (exists && tx.passThrough === false) {
      const old = txs.find((t) => t.id === tx.id);
      if (old?.groupId) {
        const others = nextTxs.filter((t) => t.id !== tx.id && t.groupId === old.groupId);
        if (others.length <= 1) {
          // grupo de 1 ya no tiene sentido: quitar passThrough y groupId al solitario
          nextTxs = nextTxs.map((t) =>
            t.groupId === old.groupId && t.id !== tx.id
              ? { ...t, passThrough: false, groupId: null }
              : t
          );
        }
      }
    }

    saveTxs(nextTxs);
    if (learnedCats) saveConfig({ ...config, categories: learnedCats });
    setAdding(false);
    setEditingTx(null);
    showToast(exists ? "Movimiento actualizado" : `${tx.type === "income" ? "Ingreso" : "Gasto"} registrado`);
  };

  const saveManyTxs = (newTxs, learnedCats) => {
    saveTxs([...newTxs, ...txs]);
    if (learnedCats) saveConfig({ ...config, categories: learnedCats });
    setImportOpen(false);
    showToast(`${newTxs.length} movimiento${newTxs.length === 1 ? "" : "s"} importado${newTxs.length === 1 ? "" : "s"}`);
  };

  const overlayOpen = chatOpen || adding || !!editingTx || accountsOpen || importOpen || excelOpen || addMenuOpen || recurringOpen || settingsOpen;

  return (
    <div>
      <StickyHeader config={config} saveConfig={saveConfig} balance={balance} dateRange={dateRange} onOpenRange={() => setRangeOpen(true)} onOpenSettings={() => setSettingsOpen(true)} />

      <div className="cc-wrap">
        {tab === "inicio" && <Dashboard config={config} txs={txs} balance={balance} dateRange={dateRange} onEdit={setEditingTx} onAddAccount={() => setAccountsOpen(true)} saveConfig={saveConfig} />}
        {tab === "movs" && <Movimientos config={config} txs={txs} dateRange={dateRange} saveTxs={saveTxs} showToast={showToast} onEdit={setEditingTx} />}
        {tab === "cats" && <Categorias config={config} txs={txs} dateRange={dateRange} saveConfig={saveConfig} showToast={showToast} saveRecurring={saveRecurring} />}
        {tab === "stats" && <Estadisticas config={config} txs={txs} dateRange={dateRange} onEdit={setEditingTx} saveConfig={saveConfig} />}
      </div>

      <BottomNav
        tab={tab}
        setTab={setTab}
        onOpenAssistant={(opts) => { setChatVoice(!!(opts && opts.voice)); setChatOpen(true); }}
        hidden={overlayOpen}
      />

      <TopFab
        open={addMenuOpen}
        onToggle={() => setAddMenuOpen((v) => !v)}
        onPickExcel={() => { setAddMenuOpen(false); setExcelOpen(true); }}
        onPickScreenshot={() => { setAddMenuOpen(false); setImportOpen(true); }}
        onPickManual={() => { setAddMenuOpen(false); setAdding(true); }}
        onPickRecurring={() => { setAddMenuOpen(false); setRecurringOpen(true); }}
        hidden={chatOpen || adding || !!editingTx || accountsOpen || importOpen || excelOpen || recurringOpen}
      />

      {(adding || editingTx) && (
        <AddModal
          config={config}
          tx={editingTx}
          txs={txs}
          onClose={() => { setAdding(false); setEditingTx(null); }}
          onSave={upsertTx}
        />
      )}

      {importOpen && (
        <ImportModal
          config={config}
          txs={txs}
          onClose={() => setImportOpen(false)}
          onSave={saveManyTxs}
        />
      )}

      {chatOpen && (
        <Assistant
          config={config}
          txs={txs}
          saveConfig={saveConfig}
          saveTxs={saveTxs}
          autoVoice={chatVoice}
          onClose={() => { setChatOpen(false); setChatVoice(false); }}
          onOpenImport={() => setImportOpen(true)}
        />
      )}

      {accountsOpen && (
        <AccountsModal
          config={config}
          txs={txs}
          saveConfig={saveConfig}
          showToast={showToast}
          resetAll={resetAll}
          onClose={() => setAccountsOpen(false)}
        />
      )}

      {excelOpen && (
        <ExcelImportModal
          config={config}
          txs={txs}
          onClose={() => setExcelOpen(false)}
          onSave={saveManyTxs}
        />
      )}

      {recurringOpen && (
        <RecurringModal
          config={config}
          onClose={() => setRecurringOpen(false)}
          onSave={saveRecurring}
        />
      )}

      {settingsOpen && (
        <SettingsModal
          config={config}
          saveConfig={saveConfig}
          onClose={() => setSettingsOpen(false)}
          showToast={showToast}
          resetAll={resetAll}
        />
      )}

      {rangeOpen && (
        <DateRangeModal
          dateRange={dateRange}
          onClose={() => setRangeOpen(false)}
          onSave={(r) => { setDateRange(r); setRangeOpen(false); }}
        />
      )}
    </div>
  );
}

/* Iconos lineales del bottom nav (outline, sin emojis) */
const NavIconHome = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 11l9-8 9 8" />
    <path d="M5 10v10h14V10" />
  </svg>
);
const NavIconHistory = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 12a9 9 0 1 0 3-6.7" />
    <path d="M3 4v4h4" />
    <path d="M12 8v4l3 2" />
  </svg>
);
const NavIconCategories = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="3" width="7" height="7" rx="1.5" />
    <rect x="14" y="3" width="7" height="7" rx="1.5" />
    <rect x="3" y="14" width="7" height="7" rx="1.5" />
    <rect x="14" y="14" width="7" height="7" rx="1.5" />
  </svg>
);
const NavIconStats = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M4 20V10" />
    <path d="M12 20V4" />
    <path d="M20 20v-6" />
  </svg>
);

/* Iconos lineales para chips de cabecera (sin emojis) */
const IconGear = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="3" />
    <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.6 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.6a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
  </svg>
);
const IconCalendar = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="4" width="18" height="18" rx="2" />
    <path d="M16 2v4M8 2v4M3 10h18" />
  </svg>
);

/* StickyHeader: header pegajoso con logo, balance, chip de rango y tabs */
/* Botón flotante del asistente — se monta en document.body para evitar
   problemas de stacking context (ancestros con transform/backdrop-filter
   rompen position:fixed). DOM nativo, no usa react-dom/createPortal. */
/* BottomNav: barra inferior con Home / History / Orb IA / Categories / Statistics */
function BottomNav({ tab, setTab, onOpenAssistant, hidden }) {
  const orbRef = useRef(null);
  const holdTimer = useRef(null);
  const heldRef = useRef(false);

  const orbBurst = () => {
    const orb = orbRef.current;
    if (!orb) return;
    const rect = orb.getBoundingClientRect();
    for (let i = 0; i < 6; i++) {
      const p = document.createElement("span");
      p.style.cssText = `position:fixed;width:4px;height:4px;border-radius:50%;background:var(--orb-mint);pointer-events:none;left:${rect.left + rect.width / 2}px;top:${rect.top + rect.height / 2}px;z-index:9999;`;
      document.body.appendChild(p);
      const angle = (i / 6) * Math.PI * 2;
      const dx = Math.cos(angle) * 22;
      const dy = Math.sin(angle) * 22;
      const anim = p.animate(
        [{ transform: "translate(0,0) scale(1)", opacity: 1 },
         { transform: `translate(${dx}px,${dy}px) scale(0)`, opacity: 0 }],
        { duration: 500, easing: "cubic-bezier(.2,.7,.2,1)" }
      );
      anim.onfinish = () => p.remove();
    }
    orb.animate(
      [{ transform: "translate(-50%,-50%) scale(1)" },
       { transform: "translate(-50%,-50%) scale(.8)", offset: .3 },
       { transform: "translate(-50%,-50%) scale(1.12)", offset: .6 },
       { transform: "translate(-50%,-50%) scale(1)" }],
      { duration: 550, easing: "cubic-bezier(.34,1.56,.64,1)" }
    );
  };

  // toque corto: abre asistente · mantener presionado: abre y dicta por voz
  const onOrbDown = () => {
    heldRef.current = false;
    if (!voiceSupported) return;
    holdTimer.current = setTimeout(() => {
      heldRef.current = true;
      if (navigator.vibrate) navigator.vibrate(12);
      orbBurst();
      onOpenAssistant({ voice: true });
    }, 450);
  };
  const onOrbUp = () => {
    if (holdTimer.current) { clearTimeout(holdTimer.current); holdTimer.current = null; }
  };
  const handleOrbClick = () => {
    if (heldRef.current) { heldRef.current = false; return; } // ya se abrió por hold
    orbBurst();
    setTimeout(() => onOpenAssistant({ voice: false }), 120);
  };

  const NAV_ITEMS = [
    ["inicio", "home",    NavIconHome],
    ["movs",   "history", NavIconHistory],
  ];
  const NAV_ITEMS_2 = [
    ["cats",  "categories", NavIconCategories],
    ["stats", "statistics", NavIconStats],
  ];

  return (
    <div className="cc-bottomnav" style={{
      opacity: hidden ? 0 : 1,
      transform: hidden ? "translateY(16px)" : "translateY(0)",
      pointerEvents: hidden ? "none" : "auto",
      transition: "opacity .2s, transform .25s cubic-bezier(.2,.7,.2,1)",
    }}>
      <div className="cc-bottomnav-inner">
        {NAV_ITEMS.map(([k, label, Icon]) => (
          <button key={k} className={`cc-nav-item ${tab === k ? "on" : ""}`} onClick={() => setTab(k)}>
            <span className="cc-nav-icon"><Icon /></span>
            <span className="cc-nav-label">{t(label)}</span>
          </button>
        ))}

        <div className="cc-orb-slot">
          <button ref={orbRef} className="cc-orb-btn"
            onClick={handleOrbClick}
            onMouseDown={onOrbDown} onMouseUp={onOrbUp} onMouseLeave={onOrbUp}
            onTouchStart={onOrbDown} onTouchEnd={onOrbUp}
            aria-label="Asistente Zafi (mantén presionado para hablar)">
            <span className="cc-orb" />
          </button>
        </div>

        {NAV_ITEMS_2.map(([k, label, Icon]) => (
          <button key={k} className={`cc-nav-item ${tab === k ? "on" : ""}`} onClick={() => setTab(k)}>
            <span className="cc-nav-icon"><Icon /></span>
            <span className="cc-nav-label">{t(label)}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

/* TopFab: botón circular (+) arriba a la derecha; abre hoja de captura */
function TopFab({ open, onToggle, onPickExcel, onPickScreenshot, onPickManual, onPickRecurring, hidden }) {
  const items = [
    { icon: "✏️", label: t("manualCapture"), desc: t("manualCaptureDesc"), onClick: onPickManual },
    { icon: "🔁", label: t("recurringMovement"), desc: t("recurringDesc"), onClick: onPickRecurring },
    { icon: "📸", label: t("fromScreenshot"), desc: t("fromScreenshotDesc"), onClick: onPickScreenshot },
    { icon: "📊", label: t("fromExcel"), desc: t("fromExcelDesc"), onClick: onPickExcel },
  ];
  return (
    <>
      <button className={`cc-fab-top ${open ? "open" : ""}`}
        onClick={onToggle}
        style={{
          opacity: hidden ? 0 : 1,
          transform: hidden ? `${open ? "rotate(45deg) " : ""}translateY(-16px)` : (open ? "rotate(45deg)" : "none"),
          pointerEvents: hidden ? "none" : "auto",
          transition: "opacity .2s, transform .25s cubic-bezier(.2,.7,.2,1)",
        }}
        aria-label="Nueva transacción">＋</button>

      {open && !hidden && (
        <div className="cc-overlay" onClick={onToggle}>
          <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
            <div className="cc-grip" />
            <div className="cc-sheet-top">
              <h2>{t("addMovement")}</h2>
              <button className="cc-sheet-close" onClick={onToggle}>×</button>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 9, marginTop: 4 }}>
              {items.map((it) => (
                <button key={it.label} className="cc-card" onClick={it.onClick}
                  style={{ display: "flex", alignItems: "center", gap: 14, padding: "15px 16px",
                    cursor: "pointer", textAlign: "left", width: "100%", fontFamily: "inherit" }}>
                  <div style={{ width: 44, height: 44, borderRadius: 12, background: "var(--surface)",
                    display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, flexShrink: 0 }}>
                    {it.icon}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: 15, color: "var(--ink)", letterSpacing: "-.01em" }}>{it.label}</div>
                    <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 2 }}>{it.desc}</div>
                  </div>
                  <span style={{ fontSize: 18, color: "var(--ink-faint)", flexShrink: 0 }}>›</span>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function StickyHeader({ config, saveConfig, balance, dateRange, onOpenRange, onOpenSettings }) {
  const [scrolled, setScrolled] = useState(false);
  useEffect(() => {
    const onScroll = () => setScrolled(window.scrollY > 4);
    window.addEventListener("scroll", onScroll, { passive: true });
    return () => window.removeEventListener("scroll", onScroll);
  }, []);
  // nombre del usuario — config.userName si existe, si no, derivado del correo
  const email = auth.currentUser?.email || "";
  const rawName = email.split("@")[0] || "Usuario";
  const emailName = rawName
    .replace(/[._\-]+/g, " ")
    .split(" ")
    .filter(Boolean)
    .map((w) => w[0].toUpperCase() + w.slice(1))
    .join(" ") || "Usuario";
  const displayName = config?.userName || emailName;
  const initial = displayName[0]?.toUpperCase() || "U";
  const avatarSrc = getAvatarSrc(config);

  return (
    <div className={`cc-top ${scrolled ? "scrolled" : ""}`}>
      <div className="cc-top-inner">
        {/* Logo centrado */}
        <div className="cc-zafi-wordmark">zafi</div>

        {/* Perfil: avatar + nombre + plan */}
        <div className="cc-profile-row">
          <button onClick={onOpenSettings}
            style={{ display: "flex", alignItems: "center", gap: 12, border: "none", background: "transparent",
              cursor: "pointer", padding: 0, textAlign: "left", minWidth: 0 }}>
            <div className="cc-avatar" style={{ overflow: "hidden" }}>
              {avatarSrc ? <img src={avatarSrc} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : initial}
            </div>
            <div style={{ minWidth: 0 }}>
              <div className="cc-profile-name">{displayName}</div>
              <div className="cc-profile-plan">Plan semanal</div>
            </div>
          </button>
          <div style={{ flex: 1 }} />
          <button className="cc-range-chip" onClick={onOpenRange}>
            <IconCalendar />
            <span>{rangeLabel(dateRange)}</span>
            <span className="cc-range-arrow">▼</span>
          </button>
        </div>
      </div>
    </div>
  );
}

/* ===================== AVATARES ========================================= */
const AVATAR_STYLES = [
  { id: "male-default", url: "data:image/jpeg;base64,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" },
  { id: "female-default", url: "data:image/jpeg;base64,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" },
  { id: "option-beanie", url: "data:image/jpeg;base64,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" },
  { id: "option-hoodie", url: "data:image/jpeg;base64,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" },
];

// Asigna avatar por defecto según el género del registro
function defaultAvatarForGender(gender) {
  if (gender === "male") return "male-default";
  if (gender === "female") return "female-default";
  return null; // "other" → solo inicial
}

// Resizes an image file to a small square base64 for storage
function resizeImageToBase64(file, size = 200) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = size; canvas.height = size;
        const ctx = canvas.getContext("2d");
        const s = Math.min(img.width, img.height);
        const sx = (img.width - s) / 2, sy = (img.height - s) / 2;
        ctx.drawImage(img, sx, sy, s, s, 0, 0, size, size);
        resolve(canvas.toDataURL("image/jpeg", 0.8));
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

function getAvatarSrc(config) {
  if (config.avatarData) return config.avatarData; // custom photo
  if (config.avatarId) {
    const av = AVATAR_STYLES.find(a => a.id === config.avatarId);
    return av ? av.url : null;
  }
  return null;
}

function AvatarPickerModal({ config, saveConfig, onClose, showToast }) {
  const userName = config.userName || t("user");
  const currentSrc = getAvatarSrc(config);
  const initial = userName.charAt(0).toUpperCase();
  const fileRef = useRef(null);

  const pickPreset = (id) => {
    const { avatarData, ...rest } = config;
    saveConfig({ ...rest, avatarId: id });
    showToast(t("avatarUpdated"));
    onClose();
  };

  const handlePhoto = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const base64 = await resizeImageToBase64(file);
    const { avatarId, ...rest } = config;
    saveConfig({ ...rest, avatarData: base64 });
    showToast(t("avatarUpdated"));
    onClose();
  };

  const clear = () => {
    const { avatarId, avatarData, ...rest } = config;
    saveConfig(rest);
    showToast(t("avatarRemoved"));
    onClose();
  };

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{t("chooseAvatar")}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>

        {/* Current avatar */}
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8, marginBottom: 18 }}>
          <div style={{ position: "relative" }}>
            <div style={{ width: 80, height: 80, borderRadius: "50%", overflow: "hidden",
              background: "var(--ink)", display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 32, color: "#fff", fontFamily: "'Fraunces',serif", fontWeight: 600,
              boxShadow: "0 4px 20px rgba(30,40,60,.18)" }}>
              {currentSrc ? <img src={currentSrc} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : initial}
            </div>
            {(config.avatarId || config.avatarData) && (
              <button onClick={clear}
                style={{ position: "absolute", top: -4, right: -4, width: 26, height: 26,
                  borderRadius: "50%", border: "none", background: "var(--surface)",
                  boxShadow: "var(--shadow-sm)", cursor: "pointer", fontSize: 13,
                  display: "flex", alignItems: "center", justifyContent: "center", color: "var(--ink-soft)" }}>
                🗑
              </button>
            )}
          </div>
          <div style={{ fontWeight: 600, fontSize: 18, color: "var(--ink)" }}>{userName}</div>
        </div>

        {/* Upload photo buttons */}
        <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
          <button className="cc-btn" onClick={() => fileRef.current?.click()}
            style={{ flex: 1, padding: "12px 10px", fontSize: 13, display: "flex", flexDirection: "column", alignItems: "center", gap: 6 }}>
            <span style={{ fontSize: 22 }}>🖼</span>
            {t("choosePhoto")}
          </button>
          <input ref={fileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={handlePhoto} />
          <button className="cc-btn" onClick={() => {
            const inp = document.createElement("input");
            inp.type = "file"; inp.accept = "image/*"; inp.capture = "user";
            inp.onchange = handlePhoto; inp.click();
          }}
            style={{ flex: 1, padding: "12px 10px", fontSize: 13, display: "flex", flexDirection: "column", alignItems: "center", gap: 6 }}>
            <span style={{ fontSize: 22 }}>📷</span>
            {t("takePhoto")}
          </button>
        </div>

        {/* Preset avatars grid */}
        <div style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 11, fontWeight: 700,
          color: "var(--ink-faint)", letterSpacing: ".06em", textTransform: "uppercase",
          marginBottom: 10 }}>Avatares</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          {AVATAR_STYLES.map(a => (
            <button key={a.id} onClick={() => pickPreset(a.id)}
              style={{ aspectRatio: "1", borderRadius: 18, overflow: "hidden",
                border: config.avatarId === a.id ? "3px solid var(--green)" : "2px solid var(--glass-border)",
                cursor: "pointer", padding: 0,
                backgroundImage: `url(${a.url})`, backgroundSize: "cover", backgroundPosition: "center",
                backgroundColor: "var(--surface)",
                boxShadow: config.avatarId === a.id ? "0 0 0 2px var(--green)" : "var(--shadow-sm)",
                transition: "border .15s, box-shadow .15s" }}>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ===================== MODAL: AJUSTES GENERALES ========================= */
function SettingsModal({ config, saveConfig, onClose, showToast, resetAll }) {
  const user = auth.currentUser;
  const email = user?.email || "";
  const [section, setSection] = useState("menu"); // menu | personal | lang | currency | theme | legal | data
  const [userName, setUserName] = useState(config.userName || "");
  const [phone, setPhone] = useState(config.phone || "");
  const [age, setAge] = useState(config.userAge ? String(config.userAge) : "");
  const [lang, setLang] = useState(config.language || "es");
  const [currency, setCurrency] = useState(config.currency || "MXN");
  const [confirmReset, setConfirmReset] = useState(false);
  const [confirmDeleteAccount, setConfirmDeleteAccount] = useState(false);
  const [busy, setBusy] = useState(false);
  const [avatarOpen, setAvatarOpen] = useState(false);
  const [saved, setSaved] = useState(false);

  const initial = userName ? userName.charAt(0).toUpperCase() : email.charAt(0).toUpperCase();
  const avatarSrc = getAvatarSrc(config);

  const savePersonal = () => {
    saveConfig({ ...config, userName: userName.trim(), phone: phone.trim(), userAge: Number(age) || null });
    showToast(t("infoUpdated"));
    setSaved(true);
  };
  const markDirty = () => setSaved(false);
  const saveLang = (l) => {
    setLang(l); setAppLang(l);
    saveConfig({ ...config, language: l });
    showToast(l === "es" ? "Idioma: Español" : "Language: English");
  };
  const saveCurrency = (c) => {
    setCurrency(c);
    saveConfig({ ...config, currency: c });
    showToast(`Moneda: ${c}`);
  };
  const doSignOut = async () => {
    setBusy(true);
    try { await signOut(auth); } catch (e) {}
    setBusy(false); onClose();
  };
  const doReset = () => { resetAll(); setConfirmReset(false); onClose(); };
  const doDeleteAccount = async () => {
    setBusy(true);
    try {
      // borrar datos de Firestore
      if (user) {
        try { await deleteDoc(doc(db, "users", user.uid, "data", "config")); } catch (e) {}
        try { await deleteDoc(doc(db, "users", user.uid, "data", "txs")); } catch (e) {}
        try { await deleteDoc(doc(db, "users", user.uid, "data", "profile")); } catch (e) {}
      }
      // borrar la cuenta de Firebase Auth
      if (user) await deleteUser(user);
    } catch (e) {
      showToast("Error al eliminar. Cierra sesión, vuelve a entrar e intenta de nuevo.");
    }
    setBusy(false); onClose();
  };

  // SVG icons for menu rows
  const IconPerson = () => <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="var(--ink-soft)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="8" r="4"/><path d="M5 20c0-4 3.5-7 7-7s7 3 7 7"/></svg>;
  const IconLang = () => <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="var(--ink-soft)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><path d="M2 12h20M12 2a15.3 15.3 0 014 10 15.3 15.3 0 01-4 10 15.3 15.3 0 01-4-10A15.3 15.3 0 0112 2z"/></svg>;
  const IconCoin = () => <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="var(--ink-soft)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><path d="M9 12h6M12 9v6"/></svg>;
  const IconBell = () => <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="var(--ink-soft)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9M13.73 21a2 2 0 01-3.46 0"/></svg>;
  const IconDoc = () => <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="var(--ink-soft)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="15" y2="17"/></svg>;
  const IconShield = () => <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="var(--ink-soft)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>;
  const IconTheme = () => <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="var(--ink-soft)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="5"/><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42"/></svg>;

  const curTheme = config.theme || "auto";
  const setTheme = (t) => {
    saveConfig({ ...config, theme: t });
    const labels = { light: "Tema claro", dark: "Tema oscuro", auto: "Tema automático" };
    showToast(labels[t]);
  };
  const themeLabel = { light: "Claro", dark: "Oscuro", auto: "Auto" }[curTheme];

  const ROW = (Icon, label, value, onClick) => (
    <button onClick={onClick} style={{ display: "flex", alignItems: "center", gap: 14, width: "100%",
      padding: "15px 0", background: "transparent",
      border: "none", borderBottom: "1px solid var(--line-soft)", cursor: "pointer",
      fontFamily: "inherit", textAlign: "left" }}>
      <span style={{ width: 28, display: "flex", justifyContent: "center", flexShrink: 0 }}><Icon /></span>
      <span style={{ flex: 1, fontSize: 15, fontWeight: 500, color: "var(--ink)" }}>{label}</span>
      {value && <span style={{ fontSize: 13, color: "var(--ink-soft)" }}>{value}</span>}
      <span style={{ fontSize: 16, color: "var(--ink-faint)" }}>›</span>
    </button>
  );

  const BACK = (title) => (
    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
      <button onClick={() => setSection("menu")}
        style={{ width: 34, height: 34, borderRadius: "50%", border: "none",
          background: "var(--surface)", cursor: "pointer", fontSize: 18,
          display: "flex", alignItems: "center", justifyContent: "center", color: "var(--ink)" }}>‹</button>
      <h2 style={{ fontSize: 20, fontWeight: 700, color: "var(--ink)", margin: 0 }}>{title}</h2>
    </div>
  );

  const CHIP_ROW = (options, current, onPick) => (
    <div style={{ display: "flex", gap: 8 }}>
      {options.map(([k, l]) => (
        <button key={k} onClick={() => onPick(k)}
          style={{ flex: 1, padding: "12px 8px", borderRadius: 14, cursor: "pointer",
            fontFamily: "'Montserrat', sans-serif", fontSize: 14, fontWeight: current === k ? 600 : 400,
            background: current === k ? "#5B6EE8" : "var(--surface)",
            color: current === k ? "#fff" : "var(--ink)",
            border: `1px solid ${current === k ? "#5B6EE8" : "var(--line)"}`,
            transition: "all .15s ease" }}>
          {l}
        </button>
      ))}
    </div>
  );

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />

        {section === "menu" && (
          <>
            <div className="cc-sheet-top">
              <h2>{t("settings")}</h2>
              <button className="cc-sheet-close" onClick={onClose}>×</button>
            </div>

            {/* Perfil header */}
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8, padding: "8px 0 20px" }}>
              <button onClick={() => setAvatarOpen(true)}
                style={{ position: "relative", border: "none", background: "transparent", cursor: "pointer", padding: 0 }}>
                <div className="cc-avatar" style={{ width: 72, height: 72, fontSize: 28, overflow: "hidden" }}>
                  {avatarSrc ? <img src={avatarSrc} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : initial}
                </div>
                <div style={{ position: "absolute", bottom: -2, right: -2, width: 24, height: 24, borderRadius: "50%",
                  background: "var(--ink)", color: "#fff", fontSize: 12,
                  display: "flex", alignItems: "center", justifyContent: "center",
                  boxShadow: "var(--shadow-sm)" }}>✎</div>
              </button>
              <div style={{ fontWeight: 600, fontSize: 18, color: "var(--ink)" }}>{userName || t("user")}</div>
              <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>{email}</div>
            </div>

            {/* Menu rows */}
            <div style={{ display: "flex", flexDirection: "column" }}>
              {ROW(IconPerson, t("personalInfo"), "", () => setSection("personal"))}
              {ROW(IconLang, t("language"), lang === "es" ? "Español" : "English", () => setSection("lang"))}
              {ROW(IconCoin, t("currency"), currency, () => setSection("currency"))}
              {ROW(IconBell, t("notifications"), t("comingSoon"), () => {})}
              {ROW(IconTheme, "Tema", themeLabel, () => setSection("theme"))}
              {ROW(IconDoc, "Aviso legal", "", () => setSection("legal"))}
              {ROW(IconShield, t("dataPrivacy"), "", () => setSection("data"))}
            </div>

            {/* Sign out */}
            <div style={{ marginTop: 24, marginBottom: 8 }}>
              <button className="cc-btn" onClick={doSignOut} disabled={busy}
                style={{ width: "100%", padding: 14, fontSize: 15, fontWeight: 600 }}>
                {busy ? t("signingOut") : t("signOut")}
              </button>
            </div>
            <div style={{ textAlign: "center", fontSize: 11, color: "var(--ink-faint)", padding: "6px 0 4px" }}>
              Zafi · Finanzas personales con IA · v1.0
            </div>
          </>
        )}

        {section === "personal" && (
          <>
            {BACK(t("personalInfo"))}
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              <div>
                <label className="cc-label">{t("name")}</label>
                <input className="cc-input" value={userName} onChange={(e) => { setUserName(e.target.value); markDirty(); }} placeholder="Tu nombre" />
              </div>
              <div>
                <label className="cc-label">{t("age")}</label>
                <input className="cc-input" value={age} type="text" inputMode="numeric" placeholder="00"
                  onChange={(e) => { setAge(e.target.value.replace(/[^0-9]/g, "").slice(0, 3)); markDirty(); }}
                  style={{ width: 100 }} />
              </div>
              <div>
                <label className="cc-label">{t("phone")}</label>
                <input className="cc-input" value={phone} onChange={(e) => { setPhone(e.target.value); markDirty(); }} placeholder="+52 664 123 4567" type="tel" />
              </div>
              <div>
                <label className="cc-label">{t("email")}</label>
                <input className="cc-input" value={email} disabled style={{ opacity: 0.6 }} />
              </div>
              <button className="cc-btn" onClick={saved ? undefined : savePersonal} disabled={saved}
                style={{ width: "100%", padding: 14,
                  background: saved ? "var(--green)" : "var(--green)",
                  color: "#fff", borderColor: "var(--green)",
                  opacity: saved ? 0.6 : 1, transition: "opacity .2s" }}>
                {saved ? "Listo ✓" : t("saveChanges")}
              </button>
            </div>
          </>
        )}

        {section === "lang" && (
          <>
            {BACK(t("language"))}
            <div style={{ minHeight: 200 }}>
              <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 16, lineHeight: 1.6 }}>
                {_lang === "es" ? "Selecciona el idioma de la app." : "Select the app language."}
              </div>
              {CHIP_ROW([["es", "🇲🇽 Español"], ["en", "🇺🇸 English"]], lang, saveLang)}
            </div>
          </>
        )}

        {section === "currency" && (
          <>
            {BACK(t("currency"))}
            <div style={{ minHeight: 200 }}>
              <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 16, lineHeight: 1.6 }}>
                {_lang === "es" ? "Moneda predeterminada para tus movimientos." : "Default currency for your transactions."}
              </div>
              {CHIP_ROW([["MXN", "🇲🇽 MXN"], ["USD", "🇺🇸 USD"], ["EUR", "🇪🇺 EUR"]], currency, saveCurrency)}
            </div>
          </>
        )}

        {section === "theme" && (
          <>
            {BACK("Tema")}
            <div style={{ minHeight: 200 }}>
              <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 16, lineHeight: 1.6 }}>
                Elige cómo se ve la app. "Automático" sigue la configuración de tu dispositivo.
              </div>
              {CHIP_ROW([["light", "Claro"], ["dark", "Oscuro"], ["auto", "Auto"]], curTheme, setTheme)}
            </div>
          </>
        )}

        {section === "legal" && (
          <>
            {BACK("Aviso legal")}
            <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
              <div style={{ padding: "14px 0", borderBottom: "1px solid var(--line-soft)" }}>
                <div style={{ fontSize: 14, fontWeight: 600, color: "var(--ink)", marginBottom: 4 }}>Términos y condiciones</div>
                <div style={{ fontSize: 12.5, color: "var(--ink-soft)", lineHeight: 1.6 }}>
                  Al usar Zafi aceptas que es una herramienta de finanzas personales. No somos asesores financieros. La información que registres es tu responsabilidad.
                </div>
              </div>
              <div style={{ padding: "14px 0", borderBottom: "1px solid var(--line-soft)" }}>
                <div style={{ fontSize: 14, fontWeight: 600, color: "var(--ink)", marginBottom: 4 }}>Política de privacidad</div>
                <div style={{ fontSize: 12.5, color: "var(--ink-soft)", lineHeight: 1.6 }}>
                  Tus datos financieros se almacenan de forma segura en Firebase (Google Cloud). No compartimos ni vendemos tu información personal. Tu historial de transacciones es privado y solo tú puedes acceder a él.
                </div>
              </div>
              <div style={{ padding: "14px 0" }}>
                <div style={{ fontSize: 14, fontWeight: 600, color: "var(--ink)", marginBottom: 4 }}>Contacto</div>
                <div style={{ fontSize: 12.5, color: "var(--ink-soft)", lineHeight: 1.6 }}>
                  ¿Preguntas o sugerencias? Escríbenos a zafi.fintech@gmail.com
                </div>
              </div>
            </div>
          </>
        )}

        {section === "data" && (
          <>
            {BACK(t("dataPrivacy"))}
            <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
              <div style={{ padding: "14px 0", borderBottom: "1px solid var(--line-soft)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 14, fontWeight: 500, color: "var(--ink)" }}>{t("exportData")}</span>
                <span style={{ fontSize: 13, color: "var(--ink-soft)" }}>{t("comingSoon")}</span>
              </div>
              {!confirmReset ? (
                <button onClick={() => setConfirmReset(true)}
                  style={{ padding: "14px 0", background: "transparent", border: "none",
                    borderBottom: "1px solid var(--line-soft)", cursor: "pointer",
                    fontFamily: "inherit", textAlign: "left", width: "100%" }}>
                  <span style={{ fontSize: 14, fontWeight: 500, color: "var(--coral)" }}>{t("resetApp")}</span>
                </button>
              ) : (
                <div style={{ padding: "14px 0", borderBottom: "1px solid var(--line-soft)" }}>
                  <div style={{ fontSize: 13, color: "var(--coral)", fontWeight: 600, marginBottom: 10 }}>
                    {t("resetConfirm")}
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button className="cc-btn" style={{ flex: 1, padding: "10px 12px", fontSize: 13 }}
                      onClick={() => setConfirmReset(false)}>{t("cancel")}</button>
                    <button className="cc-btn" style={{ flex: 1, padding: "10px 12px", fontSize: 13,
                      background: "var(--coral)", color: "#fff", borderColor: "var(--coral)" }}
                      onClick={doReset}>{t("yesDeleteAll")}</button>
                  </div>
                </div>
              )}
              {/* Eliminar cuenta */}
              {!confirmDeleteAccount ? (
                <button onClick={() => setConfirmDeleteAccount(true)}
                  style={{ padding: "14px 0", background: "transparent", border: "none",
                    cursor: "pointer", fontFamily: "inherit", textAlign: "left", width: "100%" }}>
                  <span style={{ fontSize: 14, fontWeight: 500, color: "var(--coral)" }}>Eliminar cuenta permanentemente</span>
                  <div style={{ fontSize: 12, color: "var(--ink-faint)", marginTop: 2 }}>
                    Se borrarán todos tus datos y tu cuenta de forma irreversible.
                  </div>
                </button>
              ) : (
                <div style={{ padding: "14px 0" }}>
                  <div style={{ fontSize: 13, color: "var(--coral)", fontWeight: 600, marginBottom: 10 }}>
                    ⚠️ Esta acción es IRREVERSIBLE. Se eliminará tu cuenta, todos tus movimientos, categorías y configuración.
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button className="cc-btn" style={{ flex: 1, padding: "10px 12px", fontSize: 13 }}
                      onClick={() => setConfirmDeleteAccount(false)}>{t("cancel")}</button>
                    <button className="cc-btn" style={{ flex: 1, padding: "10px 12px", fontSize: 13,
                      background: "var(--coral)", color: "#fff", borderColor: "var(--coral)" }}
                      disabled={busy}
                      onClick={doDeleteAccount}>{busy ? "Eliminando…" : "Sí, eliminar mi cuenta"}</button>
                  </div>
                </div>
              )}
            </div>
          </>
        )}

        {avatarOpen && (
          <AvatarPickerModal config={config} saveConfig={saveConfig} onClose={() => setAvatarOpen(false)} showToast={showToast} />
        )}
      </div>
    </div>
  );
}

/* ProfileNameModal: editar el nombre que se muestra en el header */
function ProfileNameModal({ current, onClose, onSave }) {
  const [name, setName] = useState(current || "");
  return (
    <div className="cc-overlay" onClick={onClose} style={{ alignItems: "center" }}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}
        style={{ borderRadius: 24, maxWidth: 360, width: "calc(100% - 40px)" }}>
        <div className="cc-grip" />
        <h2 className="cc-serif" style={{ fontSize: 20, fontWeight: 600, marginBottom: 14 }}>Tu nombre</h2>
        <input className="cc-input" placeholder="Ej. Luis Ángel" value={name}
          onChange={(e) => setName(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && name.trim() && onSave(name.trim())}
          style={{ marginBottom: 14 }} autoFocus />
        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 13, fontSize: 14 }}
          disabled={!name.trim()} onClick={() => onSave(name.trim())}>
          Guardar
        </button>
      </div>
    </div>
  );
}

/* DateRangeModal: elegir el rango global de toda la app */
function DateRangeModal({ dateRange, onClose, onSave }) {
  const r = dateRange || DEFAULT_RANGE;
  const [preset, setPreset] = useState(r.preset);
  const resolved = resolveRange(r);
  const [from, setFrom] = useState(r.from || resolved.from);
  const [to, setTo] = useState(r.to || resolved.to);

  const PRESETS = [
    { id: "today",      label: "Hoy",              emoji: "⚡" },
    { id: "week",       label: "Esta semana",      emoji: "📅" },
    { id: "month",      label: "Este mes",         emoji: "📆" },
    { id: "last-month", label: "Mes pasado",       emoji: "📆" },
    { id: "3m",         label: "Últimos 3 meses",  emoji: "📊" },
    { id: "6m",         label: "Últimos 6 meses",  emoji: "📊" },
    { id: "year",       label: "Este año",         emoji: "🗓️" },
    { id: "last-year",  label: "Año pasado",       emoji: "🗓️" },
    { id: "all",        label: "Todo el historial", emoji: "⏳" },
  ];

  function apply() {
    if (preset === "custom") {
      if (!from || !to) return;
      const f = from > to ? to : from;
      const t = from > to ? from : to;
      onSave({ preset: "custom", from: f, to: t });
    } else {
      onSave({ preset });
    }
  }

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Rango de tiempo</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 14 }}>
          Esta selección se aplica a todas las pestañas: Inicio, Movimientos, Categorías y Estadísticas.
        </p>

        <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 14 }}>
          {PRESETS.map((p) => (
            <button key={p.id} onClick={() => setPreset(p.id)}
              style={{ display: "flex", alignItems: "center", gap: 11, padding: "11px 13px",
                background: preset === p.id ? "var(--green-soft)" : "var(--surface)",
                border: `1px solid ${preset === p.id ? "var(--green)" : "var(--line)"}`,
                borderRadius: 11, cursor: "pointer", fontFamily: "inherit", fontSize: 14,
                fontWeight: preset === p.id ? 700 : 500, textAlign: "left", color: "var(--ink)" }}>
              <span style={{ fontSize: 17 }}>{p.emoji}</span>
              <span style={{ flex: 1 }}>{p.label}</span>
              {preset === p.id && <span style={{ color: "var(--green)", fontSize: 17, fontWeight: 700 }}>✓</span>}
            </button>
          ))}

          <button onClick={() => setPreset("custom")}
            style={{ display: "flex", alignItems: "center", gap: 11, padding: "11px 13px",
              background: preset === "custom" ? "var(--green-soft)" : "var(--surface)",
              border: `1px solid ${preset === "custom" ? "var(--green)" : "var(--line)"}`,
              borderRadius: 11, cursor: "pointer", fontFamily: "inherit", fontSize: 14,
              fontWeight: preset === "custom" ? 700 : 500, textAlign: "left", color: "var(--ink)",
              marginTop: 4 }}>
            <span style={{ fontSize: 17 }}>✏️</span>
            <span style={{ flex: 1 }}>Personalizado</span>
            {preset === "custom" && <span style={{ color: "var(--green)", fontSize: 17, fontWeight: 700 }}>✓</span>}
          </button>

          {preset === "custom" && (
            <div style={{ display: "flex", gap: 8, marginTop: 4, padding: "12px 13px",
              background: "var(--surface-2)", borderRadius: 11 }}>
              <div style={{ flex: 1 }}>
                <label className="cc-label" style={{ marginBottom: 4 }}>Desde</label>
                <input className="cc-input" type="date" value={from}
                  onChange={(e) => setFrom(e.target.value)} style={{ fontSize: 13 }} />
              </div>
              <div style={{ flex: 1 }}>
                <label className="cc-label" style={{ marginBottom: 4 }}>Hasta</label>
                <input className="cc-input" type="date" value={to}
                  onChange={(e) => setTo(e.target.value)} style={{ fontSize: 13 }} />
              </div>
            </div>
          )}
        </div>

        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 13, fontSize: 14 }}
          onClick={apply}>
          Aplicar
        </button>
      </div>
    </div>
  );
}

/* secciones disponibles del inicio */
const DEFAULT_SECTIONS = [
  { id: "balance", label: "Saldo destacado", icon: "💰", on: false },
  { id: "kpis", label: "Ingresos y gastos del mes", icon: "📊", on: true },
  { id: "byCategory", label: "Gastos por categoría", icon: "🏷️", on: true },
  { id: "trend", label: "Mini gráfica de saldo (30d)", icon: "📈", on: true },
  { id: "topExpenses", label: "Gastos más grandes del mes", icon: "💸", on: false },
  { id: "recent", label: "Movimientos recientes", icon: "🕐", on: true },
];

function loadSections(config) {
  const saved = config.homeSections;
  if (!Array.isArray(saved) || !saved.length) return DEFAULT_SECTIONS;
  // sincronizar con DEFAULT por si agregamos nuevas
  const known = new Set(DEFAULT_SECTIONS.map((s) => s.id));
  const cleaned = saved.filter((s) => known.has(s.id));
  DEFAULT_SECTIONS.forEach((d) => {
    if (!cleaned.find((s) => s.id === d.id)) cleaned.push({ ...d });
  });
  return cleaned;
}

/* ============================= DASHBOARD ================================= */
function Dashboard({ config, txs, balance, dateRange, onEdit, onAddAccount, saveConfig }) {
  const [view, setView] = useState("all"); // "all" o accountId
  const [configuring, setConfiguring] = useState(false);
  const sections = loadSections(config);

  const scopedTxs = view === "all" ? txs : txs.filter((t) => t.accountId === view);
  // movimientos del rango global (en lugar de "mes actual")
  const rangeTxs = txsInRange(scopedTxs, dateRange);
  const monthStat = statTxs(rangeTxs).all;
  const inc = monthStat.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const exp = monthStat.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);

  const headerBalance = view === "all" ? balance : accountBalance(config, txs, view);
  const accName = view === "all" ? "General" : (config.accounts.find((a) => a.id === view) || {}).name || "";
  const headerLabel = view === "all" ? "Balance total" : `Saldo · ${accName}`;

  const byCat = {};
  monthStat.filter((t) => t.type === "expense" && t.categoryId).forEach((t) => {
    byCat[t.categoryId] = (byCat[t.categoryId] || 0) + t.amount;
  });
  const rows = Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, 6);
  const maxCat = rows.length ? rows[0][1] : 1;
  const catOf = (id) => config.categories.find((c) => c.id === id);

  // gastos más grandes del rango (excluyendo pass-through)
  const topExpenses = rangeTxs.filter((t) => t.type === "expense" && !t.passThrough)
    .sort((a, b) => b.amount - a.amount).slice(0, 5);

  // mini-gráfica de saldo (30 días)
  const trendPoints = (() => {
    const initial = view === "all"
      ? config.accounts.reduce((s, a) => s + (a.initialBalance || 0), 0)
      : (config.accounts.find((a) => a.id === view)?.initialBalance || 0);
    const start = new Date(); start.setDate(start.getDate() - 29);
    const startK = start.toISOString().slice(0, 10);
    const sorted = [...scopedTxs].sort((a, b) => a.date.localeCompare(b.date));
    let running = initial + sorted.filter((t) => t.date < startK).reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
    const map = new Map(); map.set(startK, running);
    sorted.forEach((t) => {
      if (t.date < startK) return;
      running += t.type === "income" ? t.amount : -t.amount;
      map.set(t.date, running);
    });
    const pts = [], todayD = new Date(today());
    let last = initial + sorted.filter((t) => t.date < startK).reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
    for (let d = new Date(start); d <= todayD; d.setDate(d.getDate() + 1)) {
      const k = d.toISOString().slice(0, 10);
      if (map.has(k)) last = map.get(k);
      pts.push({ date: k, val: last });
    }
    return pts;
  })();

  const isOn = (id) => sections.find((s) => s.id === id)?.on;

  // gráfica de líneas: gasto acumulado por día en el periodo
  const expenseLinePoints = (() => {
    const expTxs = rangeTxs.filter(t => t.type === "expense");
    if (expTxs.length < 2) return [];
    const daily = {};
    expTxs.forEach(t => { daily[t.date] = (daily[t.date] || 0) + t.amount; });
    const dates = Object.keys(daily).sort();
    if (dates.length < 2) return [];
    // llenar días sin gastos entre min y max
    const start = new Date(dates[0]);
    const end = new Date(dates[dates.length - 1]);
    const pts = [];
    let cumulative = 0;
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      const k = d.toISOString().slice(0, 10);
      cumulative += daily[k] || 0;
      pts.push({ date: k, val: cumulative });
    }
    return pts;
  })();

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      {/* === selector de cuentas mejorado === */}
      {config.accounts.length > 0 && (
        <div className="cc-fade">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
            <div className="cc-label">Tus cuentas</div>
            <button className="cc-gear" onClick={() => setConfiguring(true)}><IconGear /> Personalizar</button>
          </div>
          <div className="cc-scroll-x">
            {/* tarjeta General */}
            <button className={`cc-acc-card ${view === "all" ? "on" : ""}`} onClick={() => setView("all")}>
              <div className="cc-acc-icon">∑</div>
              <div className="cc-acc-label">Total</div>
              <div className="cc-acc-name">General</div>
              <div className="cc-acc-bal cc-num" style={{ color: balance < 0 ? "var(--coral)" : "var(--ink)" }}>
                {fmt(balance)}
              </div>
              <div className="cc-acc-sub">{config.accounts.length} cuenta{config.accounts.length === 1 ? "" : "s"}</div>
            </button>
            {/* tarjetas por cuenta */}
            {config.accounts.map((a) => {
              const b = accountBalance(config, txs, a.id);
              const accTxs = txs.filter((t) => t.accountId === a.id);
              const rangeAccStat = statTxs(txsInRange(accTxs, dateRange)).all;
              const rangeFlow = rangeAccStat.reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
              const active = view === a.id;
              return (
                <button key={a.id} className={`cc-acc-card ${active ? "on" : ""}`} onClick={() => setView(a.id)}>
                  <div className="cc-acc-icon">🏦</div>
                  <div className="cc-acc-label">Cuenta</div>
                  <div className="cc-acc-name">{a.name}</div>
                  <div className="cc-acc-bal cc-num" style={{ color: b < 0 ? "var(--coral)" : "var(--ink)" }}>
                    {fmt(b)}
                  </div>
                  <div className="cc-acc-sub" style={{ color: rangeFlow < 0 ? "var(--coral)" : undefined }}>
                    {rangeFlow >= 0 ? "▲" : "▼"} {fmt(Math.abs(rangeFlow))} en el periodo
                  </div>
                </button>
              );
            })}
            {/* tarjeta agregar */}
            <button className="cc-acc-card" onClick={onAddAccount}
              style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
                minWidth: 130, borderStyle: "dashed", color: "var(--ink-soft)" }}>
              <div style={{ fontSize: 30, lineHeight: 1, marginBottom: 6, color: "var(--gold)" }}>＋</div>
              <div style={{ fontWeight: 600, fontSize: 13 }}>Agregar cuenta</div>
            </button>
          </div>
        </div>
      )}

      {/* === secciones según orden y on/off === */}
      {sections.filter((s) => s.on).map((s, idx) => {
        const delay = `${idx * 60}ms`;

        if (s.id === "balance") return (
          <div key={s.id} className="cc-card" style={{ padding: "22px 22px" }}>
            <div className="cc-label">{headerLabel}</div>
            <div className="cc-serif cc-num" style={{ fontSize: 44, fontWeight: 600, letterSpacing: "-.02em", color: headerBalance < 0 ? "var(--coral)" : "var(--ink)" }}>
              {fmt(headerBalance)}
            </div>
          </div>
        );

        if (s.id === "kpis") return (
          <div key={s.id} style={{ display: "flex", gap: 10 }} className="cc-fade">
            <div className="cc-card" style={{ flex: 1, padding: 14 }}>
              <div className="cc-label">Ingresos · {rangeLabel(dateRange)}</div>
              <div className="cc-serif cc-num" style={{ fontSize: 21, fontWeight: 500, color: "var(--ink)" }}>{fmt(inc)}</div>
            </div>
            <div className="cc-card" style={{ flex: 1, padding: 14 }}>
              <div className="cc-label">Gastos · {rangeLabel(dateRange)}</div>
              <div className="cc-serif cc-num" style={{ fontSize: 21, fontWeight: 500, color: "var(--coral)" }}>{fmt(exp)}</div>
            </div>
          </div>
        );

        if (s.id === "byCategory") return (
          <div key={s.id} className="cc-card" style={{ padding: "6px 18px" }}>
            <div className="cc-label" style={{ marginTop: 12, marginBottom: 6 }}>Gastos por categoría · {rangeLabel(dateRange)}</div>
            {rows.length === 0 ? (
              <div style={{ color: "var(--ink-soft)", fontSize: 13, padding: "8px 0 14px" }}>
                No hay gastos en el periodo.
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column" }}>
                {rows.map(([id, amt]) => {
                  const c = catOf(id);
                  const pct = maxCat ? Math.round((amt / maxCat) * 100) : 0;
                  return (
                    <div key={id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "11px 0",
                      borderBottom: "1px solid var(--line-soft)" }}>
                      <div style={{ width: 34, height: 34, borderRadius: 10, background: "var(--surface)",
                        display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>
                        {c ? c.emoji : "❔"}
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 600, fontSize: 13, color: "var(--ink)", letterSpacing: "-.01em" }}>
                          {c ? c.name : "Sin categoría"}
                        </div>
                        <div style={{ height: 5, background: "var(--surface-2)", borderRadius: 99, overflow: "hidden", marginTop: 5 }}>
                          <div style={{ height: "100%", width: `${pct}%`, background: "var(--bar-fill)", borderRadius: 99 }} />
                        </div>
                      </div>
                      <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 14, color: "var(--coral)", whiteSpace: "nowrap" }}>
                        {fmtBare(amt)}<span style={{ fontSize: 10, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
            {/* Gráfica de líneas de gastos acumulados */}
            {expenseLinePoints.length >= 2 && (
              <div style={{ marginTop: 14, paddingTop: 10, borderTop: "1px solid var(--line-soft)" }}>
                <div className="cc-label" style={{ marginBottom: 6 }}>Gasto acumulado</div>
                <LineChart points={expenseLinePoints} color="var(--coral)" />
              </div>
            )}
          </div>
        );

        if (s.id === "trend") return (
          <div key={s.id} className="cc-card" style={{ padding: 20 }}>
            <div className="cc-label" style={{ marginBottom: 8 }}>Saldo · últimos 30 días</div>
            {trendPoints.length < 2 ? (
              <div style={{ color: "var(--ink-soft)", fontSize: 14 }}>Datos insuficientes.</div>
            ) : (
              <LineChart points={trendPoints} />
            )}
          </div>
        );

        if (s.id === "topExpenses") return (
          <div key={s.id} className="cc-card" style={{ padding: 20 }}>
            <div className="cc-label" style={{ marginBottom: 10 }}>Gastos más grandes · {rangeLabel(dateRange)}</div>
            {topExpenses.length === 0 ? (
              <div style={{ color: "var(--ink-soft)", fontSize: 14 }}>Sin gastos en el periodo.</div>
            ) : (
              topExpenses.map((t) => <TxRow key={t.id} t={t} config={config} onEdit={onEdit} />)
            )}
          </div>
        );

        if (s.id === "recent") return (
          <div key={s.id} className="cc-card" style={{ padding: 20 }}>
            <div className="cc-label" style={{ marginBottom: 10 }}>
              Movimientos recientes{view !== "all" ? ` · ${accName}` : ""}
            </div>
            {scopedTxs.length === 0 ? (
              <div style={{ color: "var(--ink-soft)", fontSize: 14 }}>Sin movimientos todavía.</div>
            ) : (
              scopedTxs.slice(0, 5).map((t) => <TxRow key={t.id} t={t} config={config} onEdit={onEdit} />)
            )}
          </div>
        );

        return null;
      })}

      {configuring && (
        <HomeConfigModal
          sections={sections}
          onClose={() => setConfiguring(false)}
          onSave={(newSections) => {
            saveConfig({ ...config, homeSections: newSections });
            setConfiguring(false);
          }}
        />
      )}
    </div>
  );
}

/* ============= MODAL: PERSONALIZAR INICIO ============================== */
function HomeConfigModal({ sections, onClose, onSave }) {
  const [items, setItems] = useState(sections);
  const [dragIdx, setDragIdx] = useState(null);
  const [overIdx, setOverIdx] = useState(null);

  const toggle = (id) => {
    setItems((prev) => prev.map((s) => (s.id === id ? { ...s, on: !s.on } : s)));
  };

  // drag & drop con HTML5
  const onDragStart = (i) => (e) => {
    setDragIdx(i);
    e.dataTransfer.effectAllowed = "move";
    try { e.dataTransfer.setData("text/plain", String(i)); } catch (_) {}
  };
  const onDragOver = (i) => (e) => {
    e.preventDefault();
    if (i !== overIdx) setOverIdx(i);
  };
  const onDrop = (i) => (e) => {
    e.preventDefault();
    if (dragIdx == null || dragIdx === i) { setDragIdx(null); setOverIdx(null); return; }
    setItems((prev) => {
      const next = [...prev];
      const [moved] = next.splice(dragIdx, 1);
      next.splice(i, 0, moved);
      return next;
    });
    setDragIdx(null); setOverIdx(null);
  };
  const move = (i, dir) => {
    setItems((prev) => {
      const j = i + dir;
      if (j < 0 || j >= prev.length) return prev;
      const next = [...prev];
      [next[i], next[j]] = [next[j], next[i]];
      return next;
    });
  };
  const reset = () => setItems(DEFAULT_SECTIONS.map((s) => ({ ...s })));

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Personalizar inicio</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 16 }}>
          Activa o desactiva secciones, y arrastra para reordenarlas.
        </p>

        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 18 }}>
          {items.map((s, i) => (
            <div key={s.id}
              draggable
              onDragStart={onDragStart(i)}
              onDragOver={onDragOver(i)}
              onDrop={onDrop(i)}
              onDragEnd={() => { setDragIdx(null); setOverIdx(null); }}
              className={`cc-sortable ${!s.on ? "disabled" : ""}`}
              style={{ borderColor: overIdx === i ? "var(--gold)" : "var(--line)",
                opacity: dragIdx === i ? 0.4 : (s.on ? 1 : 0.55) }}>
              <span className="cc-grip-h">⋮⋮</span>
              <span style={{ fontSize: 18 }}>{s.icon}</span>
              <span style={{ flex: 1, fontWeight: 600, fontSize: 14 }}>{s.label}</span>
              <button className="cc-btn" onClick={() => move(i, -1)} disabled={i === 0}
                style={{ padding: "4px 8px", fontSize: 12 }}>↑</button>
              <button className="cc-btn" onClick={() => move(i, 1)} disabled={i === items.length - 1}
                style={{ padding: "4px 8px", fontSize: 12 }}>↓</button>
              <label style={{ display: "inline-flex", alignItems: "center", cursor: "pointer", marginLeft: 4 }}>
                <input type="checkbox" checked={s.on} onChange={() => toggle(s.id)}
                  style={{ width: 18, height: 18, accentColor: "var(--green)" }} />
              </label>
            </div>
          ))}
        </div>

        <div style={{ display: "flex", gap: 8 }}>
          <button className="cc-btn" style={{ padding: 12 }} onClick={reset}>Restablecer</button>
          <button className="cc-btn cc-btn-green" style={{ flex: 1, padding: 13, fontSize: 14 }}
            onClick={() => onSave(items)}>
            Guardar cambios
          </button>
        </div>
      </div>
    </div>
  );
}

/* fila de transacción */
function TxRow({ t, config, onEdit, onDelete, selectable, selected, onToggle }) {
  const c = config.categories.find((x) => x.id === t.categoryId);
  const acc = config.accounts.find((a) => a.id === t.accountId);
  const multi = config.accounts.length > 1;

  // modo selección: toda la fila es checkbox
  if (selectable) {
    return (
      <div
        onClick={() => onToggle && onToggle(t.id)}
        style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 0",
          borderBottom: "1px solid var(--line)", cursor: "pointer",
          background: selected ? "var(--green-soft)" : "transparent",
          margin: "0 -10px", paddingLeft: 10, paddingRight: 10, borderRadius: selected ? 8 : 0 }}>
        <input type="checkbox" checked={!!selected} readOnly
          style={{ width: 19, height: 19, accentColor: "var(--green)" }} />
        <div className="cc-emoji" style={{ fontSize: 22, width: 28, textAlign: "center" }}>{c ? c.emoji : "❔"}</div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontWeight: 600, fontSize: 14.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
            {t.description || (c ? c.name : "Movimiento")}
          </div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>
            {c ? c.name : "Sin categoría"} · {t.date}{multi && acc ? ` · ${acc.name}` : ""}
          </div>
        </div>
        <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 15, color: t.type === "income" ? "var(--green)" : "var(--coral)", whiteSpace: "nowrap" }}>
          {t.type === "income" ? "+" : "−"}{fmtBare(t.amount).replace("-", "")}
          <span style={{ fontSize: 10.5, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
        </div>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0",
      borderBottom: "1px solid var(--line-soft)" }}>
      {/* emoji pill */}
      <div
        onClick={() => onEdit && onEdit(t)}
        className="cc-emoji"
        style={{ width: 34, height: 34, borderRadius: 10, background: "var(--surface)",
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 17, flexShrink: 0, cursor: onEdit ? "pointer" : "default" }}>
        {c ? c.emoji : "❔"}
      </div>
      {/* text */}
      <div
        onClick={() => onEdit && onEdit(t)}
        style={{ flex: 1, minWidth: 0, cursor: onEdit ? "pointer" : "default" }}>
        <div style={{ fontWeight: 600, fontSize: 13, color: "var(--ink)",
          whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
          letterSpacing: "-.01em" }}>
          {t.description || (c ? c.name : "Movimiento")}
        </div>
        <div style={{ fontSize: 10.5, color: "var(--ink-soft)", marginTop: 1.5, fontWeight: 500 }}>
          {c ? c.name : "Sin categoría"}{multi && acc ? ` · ${acc.name}` : ""}
        </div>
      </div>
      {/* amount */}
      <div
        onClick={() => onEdit && onEdit(t)}
        className="cc-num"
        style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 15,
          color: t.type === "income" ? "var(--green)" : "var(--coral)",
          whiteSpace: "nowrap", letterSpacing: "-.01em",
          cursor: onEdit ? "pointer" : "default" }}>
        {t.type === "income" ? "+" : "−"}{fmtBare(t.amount).replace("-", "")}
        <span style={{ fontSize: 10.5, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
      </div>
      {/* delete × */}
      {onDelete && (
        <button className="cc-sheet-close"
          style={{ width: 28, height: 28, fontSize: 16, flexShrink: 0 }}
          onClick={(e) => { e.stopPropagation(); onDelete(t.id); }}>×</button>
      )}
    </div>
  );
}

/* agrupa la lista en [{type:'header', date, income, expense}, {type:'tx', t}, ...] */
function renderGroupedByDay(list) {
  const out = [];
  let currentDate = null, headerRef = null;
  for (const t of list) {
    if (t.date !== currentDate) {
      currentDate = t.date;
      headerRef = { type: "header", date: t.date, income: 0, expense: 0 };
      out.push(headerRef);
    }
    // los pass-through no inflan el total del día
    if (!t.passThrough) {
      if (t.type === "income") headerRef.income += t.amount;
      else headerRef.expense += t.amount;
    }
    out.push({ type: "tx", t });
  }
  return out;
}

/* ============================ MOVIMIENTOS ================================ */
function Movimientos({ config, txs, dateRange, saveTxs, showToast, onEdit }) {
  const [filter, setFilter] = useState("all");
  const [sortBy, setSortBy] = useState("date-desc"); // date-desc | date-asc | amount-desc | amount-asc | account
  const [selectMode, setSelectMode] = useState(false);
  const [selected, setSelected] = useState(new Set());
  const [confirmDelete, setConfirmDelete] = useState(null); // ids[] o null
  const [accView, setAccView] = useState("all"); // account filter
  const multi = config.accounts.length > 1;

  // filtrar por cuenta, luego por rango global, luego por tipo
  const accTxs = accView === "all" ? txs : txs.filter((t) => t.accountId === accView);
  const rangeTxs = txsInRange(accTxs, dateRange);
  const filtered = rangeTxs.filter((t) => filter === "all" || t.type === filter);
  const list = [...filtered].sort((a, b) => {
    if (sortBy === "date-desc") return b.date.localeCompare(a.date) || b.id.localeCompare(a.id);
    if (sortBy === "date-asc")  return a.date.localeCompare(b.date) || a.id.localeCompare(b.id);
    if (sortBy === "amount-desc") return b.amount - a.amount;
    if (sortBy === "amount-asc")  return a.amount - b.amount;
    if (sortBy === "account") {
      const an = (config.accounts.find((x) => x.id === a.accountId) || {}).name || "";
      const bn = (config.accounts.find((x) => x.id === b.accountId) || {}).name || "";
      return an.localeCompare(bn) || b.date.localeCompare(a.date);
    }
    return 0;
  });

  // resumen del periodo (excluyendo pass-through)
  const rangeStat = statTxs(rangeTxs).all;
  const totalIn = rangeStat.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const totalOut = rangeStat.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  // top 3 categorías del tipo filtrado (o de gastos para "Todos")
  const topCatType = filter === "income" ? "income" : "expense";
  const topByCat = {};
  rangeStat.filter((t) => t.type === topCatType && t.categoryId).forEach((t) => {
    topByCat[t.categoryId] = (topByCat[t.categoryId] || 0) + t.amount;
  });
  const topCatRows = Object.entries(topByCat).sort((a, b) => b[1] - a[1]).slice(0, 3);
  const topTotal = Object.values(topByCat).reduce((s, v) => s + v, 0);

  const toggleOne = (id) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const selectAll = () => setSelected(new Set(list.map((t) => t.id)));
  const clearSel = () => setSelected(new Set());

  const askDeleteOne = (id) => setConfirmDelete([id]);
  const askDeleteMany = () => {
    if (selected.size === 0) return;
    setConfirmDelete(Array.from(selected));
  };
  const doDelete = () => {
    if (!confirmDelete) return;
    const set = new Set(confirmDelete);
    saveTxs(txs.filter((t) => !set.has(t.id)));
    showToast(`${confirmDelete.length} movimiento${confirmDelete.length === 1 ? "" : "s"} eliminado${confirmDelete.length === 1 ? "" : "s"}`);
    setConfirmDelete(null);
    setSelected(new Set());
    setSelectMode(false);
  };

  const exitSelect = () => { setSelectMode(false); setSelected(new Set()); };

  return (
    <div>
      {/* selector de cuenta */}
      {multi && (
        <div style={{ marginBottom: 12 }}>
          <div className="cc-scroll-x">
            <button className={`cc-acc-card ${accView === "all" ? "on" : ""}`} onClick={() => setAccView("all")}
              style={{ minWidth: 100 }}>
              <div className="cc-acc-label">Todas</div>
              <div className="cc-acc-name">General</div>
            </button>
            {config.accounts.map((a) => (
              <button key={a.id} className={`cc-acc-card ${accView === a.id ? "on" : ""}`} onClick={() => setAccView(a.id)}
                style={{ minWidth: 100 }}>
                <div className="cc-acc-label">🏦</div>
                <div className="cc-acc-name">{a.name}</div>
              </button>
            ))}
          </div>
        </div>
      )}
      {/* barra superior: filtro normal o info de selección */}
      {!selectMode ? (
        <>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
            <div className="cc-tabs" style={{ flex: 1 }}>
              {[["all", "Todos"], ["income", "Ingresos"], ["expense", "Gastos"]].map(([k, l]) => (
                <button key={k} className={`cc-tab ${filter === k ? "on" : ""}`} onClick={() => setFilter(k)}>{l}</button>
              ))}
            </div>
            {list.length > 0 && (
              <button className="cc-btn" style={{ padding: "8px 12px", fontSize: 12, whiteSpace: "nowrap" }}
                onClick={() => setSelectMode(true)}>
                Seleccionar
              </button>
            )}
          </div>
          {filtered.length > 0 && (
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
              <span style={{ fontSize: 12, fontWeight: 700, color: "var(--ink-soft)", textTransform: "uppercase", letterSpacing: ".06em" }}>
                Ordenar
              </span>
              <select className="cc-select" value={sortBy} onChange={(e) => setSortBy(e.target.value)}
                style={{ flex: 1, fontSize: 13, padding: "7px 10px" }}>
                <option value="date-desc">📅 Fecha — más reciente primero</option>
                <option value="date-asc">📅 Fecha — más antiguo primero</option>
                <option value="amount-desc">💰 Monto — mayor a menor</option>
                <option value="amount-asc">💰 Monto — menor a mayor</option>
                {config.accounts.length > 1 && (
                  <option value="account">🏦 Por cuenta</option>
                )}
              </select>
              <span style={{ fontSize: 11.5, color: "var(--ink-soft)", whiteSpace: "nowrap" }}>
                {filtered.length} mov.
              </span>
            </div>
          )}
        </>
      ) : (
        <div className="cc-card" style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, padding: "9px 12px",
          borderRadius: 14 }}>
          <button className="cc-sheet-close" onClick={exitSelect}
            style={{ width: 30, height: 30, fontSize: 17, flexShrink: 0 }}>×</button>
          <span style={{ flex: 1, fontWeight: 600, fontSize: 14, color: "var(--ink)" }}>
            {selected.size} seleccionado{selected.size === 1 ? "" : "s"}
          </span>
          <button className="cc-btn" style={{ padding: "6px 11px", fontSize: 12 }}
            onClick={selected.size === list.length ? clearSel : selectAll}>
            {selected.size === list.length ? "Ninguno" : "Todos"}
          </button>
          <button className="cc-btn" style={{ padding: "6px 11px", fontSize: 12,
            background: selected.size === 0 ? "var(--surface-2)" : "var(--coral)", color: selected.size === 0 ? "var(--ink-soft)" : "#fff",
            borderColor: selected.size === 0 ? "var(--line)" : "var(--coral)" }}
            disabled={selected.size === 0} onClick={askDeleteMany}>
            🗑 Eliminar
          </button>
        </div>
      )}

      {/* tarjeta resumen del periodo */}
      {!selectMode && rangeStat.length > 0 && (
        <SummaryCard filter={filter} totalIn={totalIn} totalOut={totalOut}
          topCatRows={topCatRows} topTotal={topTotal} config={config} />
      )}

      <div className="cc-card" style={{ padding: "6px 18px" }}>
        {list.length === 0 ? (
          <div style={{ color: "var(--ink-soft)", fontSize: 14, padding: "16px 0" }}>Nada por aquí todavía.</div>
        ) : (sortBy === "date-desc" || sortBy === "date-asc") ? (
          renderGroupedByDay(list).map((entry) =>
            entry.type === "header" ? (
              <div key={`h-${entry.date}`} className="cc-day-sep">
                {(() => { const p = dayParts(entry.date); return (
                  <>
                    <span className="cc-day-num">{p.num}</span>
                    <span className="cc-day-name">{p.desc}</span>
                  </>
                ); })()}
                <div className="cc-day-totals">
                  {entry.income > 0 && <span className="pos">+{fmtBare(entry.income)}</span>}
                  {entry.expense > 0 && <span className="neg">−{fmtBare(entry.expense)}</span>}
                </div>
              </div>
            ) : (
              <TxRow key={entry.t.id} t={entry.t} config={config}
                onEdit={selectMode ? null : onEdit}
                onDelete={selectMode ? null : askDeleteOne}
                selectable={selectMode}
                selected={selected.has(entry.t.id)}
                onToggle={toggleOne}
              />
            )
          )
        ) : (
          list.map((t) => (
            <TxRow key={t.id} t={t} config={config}
              onEdit={selectMode ? null : onEdit}
              onDelete={selectMode ? null : askDeleteOne}
              selectable={selectMode}
              selected={selected.has(t.id)}
              onToggle={toggleOne}
            />
          ))
        )}
      </div>

      {list.length > 0 && !selectMode && (
        <div style={{ fontSize: 12, color: "var(--ink-soft)", textAlign: "center", marginTop: 10 }}>
          Toca un movimiento para editarlo · ✕ para eliminarlo
        </div>
      )}

      {confirmDelete && (
        <ConfirmDialog
          title={confirmDelete.length === 1 ? "¿Eliminar este movimiento?" : `¿Eliminar ${confirmDelete.length} movimientos?`}
          message="Esta acción no se puede deshacer."
          confirmLabel="Eliminar"
          danger
          onCancel={() => setConfirmDelete(null)}
          onConfirm={doDelete}
        />
      )}
    </div>
  );
}

/* diálogo de confirmación reutilizable */
function ConfirmDialog({ title, message, confirmLabel = "Confirmar", danger, onCancel, onConfirm }) {
  return (
    <div className="cc-overlay" onClick={onCancel} style={{ alignItems: "center" }}>
      <div onClick={(e) => e.stopPropagation()}
        style={{ background: "var(--bg)", borderRadius: 18, maxWidth: 400, width: "90%", padding: 22, animation: "ccUp .25s" }}>
        <h3 className="cc-serif" style={{ fontSize: 19, fontWeight: 600, marginBottom: 8 }}>{title}</h3>
        {message && <p style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 18 }}>{message}</p>}
        <div style={{ display: "flex", gap: 8 }}>
          <button className="cc-btn" style={{ flex: 1, padding: 12 }} onClick={onCancel}>Cancelar</button>
          <button
            className={danger ? "cc-btn" : "cc-btn cc-btn-green"}
            style={{ flex: 1, padding: 12,
              background: danger ? "var(--coral)" : undefined,
              color: danger ? "#fff" : undefined,
              borderColor: danger ? "var(--coral)" : undefined }}
            onClick={onConfirm}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================ CATEGORÍAS ================================= */
function Categorias({ config, txs, dateRange, saveConfig, showToast, saveRecurring }) {
  const [editing, setEditing] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null); // categoría a eliminar
  const [recurringOpen, setRecurringOpen] = useState(false);
  const [accView, setAccView] = useState("all");
  const multi = config.accounts.length > 1;
  const visibleAccounts = accView === "all" ? config.accounts : config.accounts.filter((a) => a.id === accView);

  const save = (cat) => {
    let cats;
    if (cat.id) cats = config.categories.map((c) => (c.id === cat.id ? cat : c));
    else cats = [...config.categories, { ...cat, id: uid(), keywords: [] }];
    saveConfig({ ...config, categories: cats });
    setEditing(null);
    showToast("Categoría guardada");
  };
  const askDel = (cat) => setConfirmDel(cat);
  const doDel = () => {
    if (!confirmDel) return;
    saveConfig({ ...config, categories: config.categories.filter((c) => c.id !== confirmDel.id) });
    showToast("Categoría eliminada");
    setConfirmDel(null);
  };

  // calcular totales por categoría en el rango (excluyendo pass-through)
  const rangeStat = statTxs(txsInRange(txs || [], dateRange)).all;
  const totalsByCat = {};
  rangeStat.forEach((t) => {
    if (t.categoryId) totalsByCat[t.categoryId] = (totalsByCat[t.categoryId] || 0) + t.amount;
  });

  const recurring = config.recurring || [];
  const activeRec = recurring.filter((r) => isRecActive(r));
  const accName = (id) => config.accounts.find((a) => a.id === id)?.name || "—";
  const catFor = (id) => config.categories.find((c) => c.id === id);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "0 6px" }}>
        Cada cuenta tiene sus propias categorías. Los totales son de <b>{rangeLabel(dateRange)}</b>.
      </div>

      {/* selector de cuenta */}
      {multi && (
        <div style={{ marginBottom: 4 }}>
          <div className="cc-scroll-x">
            <button className={`cc-acc-card ${accView === "all" ? "on" : ""}`} onClick={() => setAccView("all")}
              style={{ minWidth: 100 }}>
              <div className="cc-acc-label">Todas</div>
              <div className="cc-acc-name">General</div>
            </button>
            {config.accounts.map((a) => (
              <button key={a.id} className={`cc-acc-card ${accView === a.id ? "on" : ""}`} onClick={() => setAccView(a.id)}
                style={{ minWidth: 100 }}>
                <div className="cc-acc-label">🏦</div>
                <div className="cc-acc-name">{a.name}</div>
              </button>
            ))}
          </div>
        </div>
      )}

      {/* ===== Movimientos recurrentes ===== */}
      <div className="cc-card" style={{ padding: 18 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 18 }}>🔁</span>
            <span className="cc-serif" style={{ fontSize: 17, fontWeight: 600 }}>Movimientos recurrentes</span>
          </div>
          <button className="cc-btn" style={{ padding: "5px 11px", fontSize: 12 }} onClick={() => setRecurringOpen(true)}>
            {recurring.length ? "Gestionar" : "＋ Nuevo"}
          </button>
        </div>
        {recurring.length === 0 ? (
          <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "6px 0 2px" }}>
            No tienes movimientos recurrentes. Crea uno para que se registre automáticamente (renta, sueldo, suscripciones…).
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", marginTop: 6 }}>
            {recurring.map((r) => {
              const c = catFor(r.categoryId);
              return (
                <button key={r.id} onClick={() => setRecurringOpen(true)}
                  style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0",
                    background: "transparent", border: "none", borderBottom: "1px solid var(--line-soft)",
                    cursor: "pointer", fontFamily: "inherit", textAlign: "left", width: "100%",
                    opacity: isRecActive(r) ? 1 : 0.5 }}>
                  <div className="cc-emoji" style={{ width: 34, height: 34, borderRadius: 10, background: "var(--surface)",
                    display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>
                    {c ? c.emoji : "🔁"}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: 13.5, color: "var(--ink)", letterSpacing: "-.01em" }}>
                      {r.description}{!isRecActive(r) && <span style={{ fontWeight: 500, color: "var(--ink-faint)" }}> · pausado</span>}
                    </div>
                    <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>
                      {FREQ_LABELS_FN()[r.freq]} · {accName(r.accountId)}{c ? ` · ${c.name}` : ""}
                    </div>
                  </div>
                  <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 14,
                    color: r.type === "income" ? "var(--green)" : "var(--coral)", whiteSpace: "nowrap" }}>
                    {r.type === "income" ? "+" : "−"}{fmtBare(r.amount)}<span style={{ fontSize: 10, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
                  </div>
                </button>
              );
            })}
            {activeRec.length > 0 && (
              <div style={{ fontSize: 11, color: "var(--ink-faint)", marginTop: 10 }}>
                {activeRec.length} activo{activeRec.length === 1 ? "" : "s"} · se generan solos en su fecha
              </div>
            )}
          </div>
        )}
      </div>

      {visibleAccounts.map((acc) => {
        const accCats = config.categories.filter((c) => c.accountId === acc.id);
        return (
          <div key={acc.id} className="cc-card" style={{ padding: 18 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
              <span style={{ fontSize: 18 }}>🏦</span>
              <span className="cc-serif" style={{ fontSize: 17, fontWeight: 600 }}>{acc.name}</span>
            </div>
            {[["expense", "Gastos"], ["income", "Ingresos"]].map(([type, label]) => {
              const list = accCats.filter((c) => c.type === type);
              return (
                <div key={type}>
                  <div className="cc-label" style={{ marginTop: 12, marginBottom: 4 }}>{label}</div>
                  {list.length === 0 && (
                    <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "3px 0" }}>Sin categorías.</div>
                  )}
                  {list.map((c) => {
                    const total = totalsByCat[c.id] || 0;
                    return (
                      <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 0", borderBottom: "1px solid var(--line)" }}>
                        <span className="cc-emoji" style={{ fontSize: 19 }}>{c.emoji}</span>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 600, fontSize: 14 }}>{c.name}</div>
                          {total > 0 && (
                            <div className="cc-num" style={{ fontSize: 11.5, color: "var(--ink-soft)" }}>
                              {fmt(total)} en el periodo
                            </div>
                          )}
                        </div>
                        <button className="cc-btn" style={{ padding: "4px 9px", fontSize: 12 }} onClick={() => setEditing(c)}>Editar</button>
                        <button className="cc-btn" style={{ padding: "4px 8px", fontSize: 12 }} onClick={() => askDel(c)}>✕</button>
                      </div>
                    );
                  })}
                </div>
              );
            })}
            <button className="cc-btn" style={{ marginTop: 12, fontSize: 13 }}
              onClick={() => setEditing({ accountId: acc.id, type: "expense", emoji: "📦", name: "" })}>
              ＋ Agregar categoría a {acc.name}
            </button>
          </div>
        );
      })}
      {editing && <CatModal cat={editing} accounts={config.accounts} onClose={() => setEditing(null)} onSave={save} />}
      {recurringOpen && (
        <RecurringModal config={config} onClose={() => setRecurringOpen(false)} onSave={saveRecurring} />
      )}
      {confirmDel && (
        <ConfirmDialog
          title="¿Eliminar esta categoría?"
          message={<>La categoría <b>{confirmDel.emoji} {confirmDel.name}</b> se eliminará. Los movimientos asociados quedarán sin categoría (no se borran).</>}
          confirmLabel="Eliminar"
          danger
          onCancel={() => setConfirmDel(null)}
          onConfirm={doDel}
        />
      )}
    </div>
  );
}

function CatModal({ cat, accounts, onClose, onSave }) {
  const [name, setName] = useState(cat.name || "");
  const [emoji, setEmoji] = useState(cat.emoji || "📦");
  const [type, setType] = useState(cat.type || "expense");
  const [accountId, setAccountId] = useState(cat.accountId || accounts[0].id);

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{cat.id ? "Editar categoría" : "Nueva categoría"}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <div style={{ width: 76 }}>
            <label className="cc-label">Emoji</label>
            <input className="cc-input" style={{ textAlign: "center", fontSize: 22 }} value={emoji}
              onChange={(e) => setEmoji(e.target.value.slice(0, 2) || "📦")} />
          </div>
          <div style={{ flex: 1 }}>
            <label className="cc-label">Nombre</label>
            <input className="cc-input" value={name} placeholder="Ej. Mascotas" onChange={(e) => setName(e.target.value)} />
          </div>
        </div>
        {accounts.length > 1 && (
          <div style={{ marginBottom: 14 }}>
            <label className="cc-label">Cuenta</label>
            <select className="cc-select" value={accountId} onChange={(e) => setAccountId(e.target.value)}>
              {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </div>
        )}
        <label className="cc-label">Tipo</label>
        <div className="cc-tabs" style={{ marginBottom: 22 }}>
          {[["expense", "Gasto"], ["income", "Ingreso"]].map(([k, l]) => (
            <button key={k} className={`cc-tab ${type === k ? "on" : ""}`} onClick={() => setType(k)}>{l}</button>
          ))}
        </div>
        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 13 }}
          disabled={!name.trim()}
          onClick={() => onSave({ ...cat, name: name.trim(), emoji, type, accountId })}>
          Guardar
        </button>
      </div>
    </div>
  );
}

/* ================== MODAL: GESTIONAR CUENTAS ============================= */
function AccountsModal({ config, txs, saveConfig, showToast, resetAll, onClose }) {
  const [editing, setEditing] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [confirmReset, setConfirmReset] = useState(false);

  const save = (acc) => {
    let accounts, categories = config.categories;
    if (acc.id) {
      accounts = config.accounts.map((a) => (a.id === acc.id ? acc : a));
    } else {
      const nid = uid();
      accounts = [...config.accounts, { ...acc, id: nid }];
      categories = [...config.categories, ...defaultCatsForAccount(nid)];
    }
    saveConfig({ ...config, accounts, categories, accountMode: accounts.length > 1 ? "multiple" : "single" });
    setEditing(null);
    showToast("Cuenta guardada");
  };
  const askDel = (acc) => {
    if (config.accounts.length <= 1) {
      showToast("Esta es tu última cuenta. Si quieres reiniciar, usa 'Empezar desde cero' abajo");
      return;
    }
    setConfirmDel(acc);
  };
  const doDel = () => {
    if (!confirmDel) return;
    const id = confirmDel.id;
    const accounts = config.accounts.filter((a) => a.id !== id);
    const categories = config.categories.filter((c) => c.accountId !== id);
    saveConfig({ ...config, accounts, categories, accountMode: accounts.length > 1 ? "multiple" : "single" });
    showToast("Cuenta eliminada");
    setConfirmDel(null);
  };
  const doReset = async () => {
    setConfirmReset(false);
    onClose();
    if (resetAll) await resetAll();
  };

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Tus cuentas</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <div className="cc-card" style={{ padding: 14, marginBottom: 12 }}>
          {config.accounts.map((a) => {
            const b = accountBalance(config, txs, a.id);
            return (
              <div key={a.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "11px 0", borderBottom: "1px solid var(--line)" }}>
                <span style={{ fontSize: 20 }}>🏦</span>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 600, fontSize: 15 }}>{a.name}</div>
                  <div style={{ fontSize: 11.5, color: "var(--ink-soft)" }}>Saldo inicial: {fmt(a.initialBalance || 0)}</div>
                </div>
                <span className="cc-num" style={{ fontWeight: 700, color: b < 0 ? "var(--coral)" : "var(--green)" }}>{fmt(b)}</span>
                <button className="cc-btn" style={{ padding: "5px 10px", fontSize: 12 }} onClick={() => setEditing(a)}>Editar</button>
                <button className="cc-btn" style={{ padding: "5px 9px", fontSize: 12 }} onClick={() => askDel(a)}>✕</button>
              </div>
            );
          })}
          <button className="cc-btn cc-btn-primary" style={{ marginTop: 14, fontSize: 13, width: "100%" }}
            onClick={() => setEditing({ name: "", initialBalance: 0 })}>
            ＋ Agregar cuenta
          </button>
        </div>
        <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "0 6px" }}>
          El <b>saldo inicial</b> es el dinero que ya tienes en cada cuenta. El saldo de la derecha es ese monto ajustado con tus movimientos.
        </div>

        {/* Cerrar sesión */}
        <div style={{ marginTop:22, paddingTop:18, borderTop:"1px solid var(--line)" }}>
          <button onClick={() => signOut(auth)}
            style={{ width:"100%", padding:"11px 14px", fontSize:14, fontFamily:"inherit",
              background:"var(--surface-2)", color:"var(--ink-soft)",
              border:"1px solid var(--line)", borderRadius:12, cursor:"pointer",
              display:"flex", alignItems:"center", justifyContent:"center", gap:8 }}>
            🚪 Cerrar sesión
          </button>
        </div>

        {/* Zona peligrosa: empezar desde cero */}
        {resetAll && (
          <div style={{ marginTop: 22, paddingTop: 18, borderTop: "1px dashed var(--line)" }}>
            <div className="cc-label" style={{ color: "var(--coral)", marginBottom: 8 }}>Zona peligrosa</div>
            <button
              onClick={() => setConfirmReset(true)}
              style={{ width: "100%", padding: "11px 14px", fontFamily: "inherit", fontSize: 13.5, fontWeight: 600,
                background: "var(--coral-soft)", color: "var(--coral)",
                border: "1px solid var(--coral)", borderRadius: 12, cursor: "pointer",
                transition: ".2s", display: "flex", alignItems: "center", justifyContent: "center", gap: 7 }}>
              🔄 Empezar desde cero
            </button>
            <div style={{ fontSize: 11.5, color: "var(--ink-soft)", marginTop: 7, lineHeight: 1.4 }}>
              Borra todas tus cuentas, categorías y movimientos. La app vuelve al onboarding inicial como usuario nuevo. No se puede deshacer.
            </div>
          </div>
        )}

        {editing && <AccountModal acc={editing} onClose={() => setEditing(null)} onSave={save} />}
        {confirmDel && (
          <ConfirmDialog
            title="¿Eliminar esta cuenta?"
            message={<>La cuenta <b>{confirmDel.name}</b> y todas sus categorías se eliminarán. Los movimientos en esa cuenta quedarán huérfanos.</>}
            confirmLabel="Eliminar"
            danger
            onCancel={() => setConfirmDel(null)}
            onConfirm={doDel}
          />
        )}
        {confirmReset && (
          <ConfirmDialog
            title="¿Empezar desde cero?"
            message={<>Esto eliminará <b>TODAS</b> tus cuentas, categorías y movimientos para siempre. La app volverá al onboarding como si fueras un usuario nuevo. <b>Esta acción no se puede deshacer.</b></>}
            confirmLabel="Sí, borrar todo"
            danger
            onCancel={() => setConfirmReset(false)}
            onConfirm={doReset}
          />
        )}
      </div>
    </div>
  );
}

function AccountModal({ acc, onClose, onSave }) {
  const [name, setName] = useState(acc.name || "");
  const [bal, setBal] = useState(acc.initialBalance != null ? String(acc.initialBalance) : "");

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{acc.id ? "Editar cuenta" : "Nueva cuenta"}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <div style={{ marginBottom: 14 }}>
          <label className="cc-label">Nombre</label>
          <input className="cc-input" placeholder="Efectivo, BBVA, Santander, Tarjeta…" value={name}
            onChange={(e) => setName(e.target.value)} />
        </div>
        <div style={{ marginBottom: 22 }}>
          <label className="cc-label">Saldo inicial</label>
          <div className="cc-amount-display">
            <span className="cc-amount-currency">$</span>
            <input className="cc-num" type="text" inputMode="decimal" placeholder="0.00"
              value={bal}
              onChange={(e) => {
                const v = e.target.value.replace(/[^0-9.]/g, "");
                if ((v.match(/\./g) || []).length <= 1) setBal(v);
              }}
              style={{ width: `${Math.max((bal || "0.00").length, 4)}ch` }} />
            <span className="cc-amount-mxn">mxn</span>
          </div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 6, textAlign: "center" }}>
            Dinero que tienes ahorita en esta cuenta. Puede ser 0.
          </div>
        </div>
        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 13 }}
          disabled={!name.trim()}
          onClick={() => onSave({ ...acc, name: name.trim(), initialBalance: parseFloat(bal) || 0 })}>
          Guardar
        </button>
      </div>
    </div>
  );
}

/* ===================== MODAL: NUEVO MOVIMIENTO =========================== */
function AddModal({ config, tx, txs, onClose, onSave }) {
  const editing = !!tx;
  const [type, setType] = useState(tx ? tx.type : "expense");
  const [amount, setAmount] = useState(tx ? String(tx.amount) : "");
  const [desc, setDesc] = useState(tx ? tx.description : "");
  const [accountId, setAccountId] = useState(
    tx ? tx.accountId : (config.accounts.length === 1 ? config.accounts[0].id : "")
  );
  const [date, setDate] = useState(tx ? tx.date : today());
  const [catId, setCatId] = useState(tx ? tx.categoryId : "auto");
  const [phase, setPhase] = useState("form");

  const cats = config.categories.filter((c) => c.type === type && c.accountId === accountId);

  /* categorizador: keywords locales -> Claude -> preguntar */
  async function categorize() {
    const nd = norm(desc);
    // primero, claves locales de las categorías de ESTA cuenta
    for (const c of cats) {
      for (const kw of c.keywords || []) {
        if (nd.includes(norm(kw))) return { catId: c.id, sure: true };
      }
    }
    // si no hay match, revisar claves aprendidas en otras cuentas con el mismo nombre/tipo
    for (const c of cats) {
      const twins = config.categories.filter(
        (x) => x.id !== c.id && x.type === c.type && norm(x.name).trim() === norm(c.name).trim()
      );
      for (const twin of twins) {
        for (const kw of twin.keywords || []) {
          if (nd.includes(norm(kw))) return { catId: c.id, sure: true };
        }
      }
    }
    try {
      const sys =
        'Eres un clasificador de transacciones personales. Responde SOLO con un objeto JSON, sin markdown ni texto extra: {"categoryId":"<id>" o null,"confident":true|false}. Devuelve null y confident:false si no estás razonablemente seguro.';
      const user = `Tipo: ${type === "income" ? "ingreso" : "gasto"}. Descripción: "${desc}". Categorías disponibles: ${JSON.stringify(
        cats.map((c) => ({ id: c.id, name: c.name }))
      )}`;
      const raw = await callClaude(sys, [{ role: "user", content: user }]);
      const p = parseJSON(raw);
      const valid = cats.some((c) => c.id === p.categoryId);
      if (valid && p.confident) return { catId: p.categoryId, sure: true };
      if (valid) return { catId: p.categoryId, sure: false };
    } catch (e) { /* sin API: preguntar */ }
    return { catId: null, sure: false };
  }

  function finalize(finalCatId, learn) {
    let learnedCats = null;
    if (learn && finalCatId) {
      const kws = extractKW(desc).slice(0, 3);
      if (kws.length) {
        const target = config.categories.find((c) => c.id === finalCatId);
        const sameName = target ? norm(target.name).trim() : null;
        learnedCats = config.categories.map((c) =>
          sameName && norm(c.name).trim() === sameName && c.type === target.type
            ? { ...c, keywords: [...new Set([...(c.keywords || []), ...kws])] }
            : c
        );
      }
    }
    onSave(
      { id: tx ? tx.id : uid(), type, amount: Math.abs(parseFloat(amount)),
        description: desc.trim(), categoryId: finalCatId, accountId, date },
      learnedCats,
      null
    );
  }

  async function handleSave() {
    if (!amount || parseFloat(amount) <= 0) return;
    if (catId !== "auto") { finalize(catId, true); return; }
    // Si solo hay una categoría disponible, usarla directo
    if (cats.length === 1) { finalize(cats[0].id, true); return; }
    // Intentar detectar con keywords locales primero (sin IA)
    const nd = norm(desc);
    for (const c of cats) {
      for (const kw of c.keywords || []) {
        if (nd.includes(norm(kw))) { finalize(c.id, false); return; }
      }
    }
    // Si hay desc, intentar con IA — pero si falla, ir directo a preguntar
    if (desc.trim()) {
      setPhase("detecting");
      const r = await categorize();
      if (r.catId && r.sure) { finalize(r.catId, false); return; }
      if (r.catId) setCatId(r.catId);
    }
    // Sin match seguro → preguntar al usuario
    setPhase("ask");
  }

  /* pantalla: preguntar categoría */
  if (phase === "ask") {
    return (
      <div className="cc-overlay" onClick={onClose}>
        <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
          <div className="cc-grip" />
          <h2 className="cc-serif" style={{ fontSize: 21, fontWeight: 600, marginBottom: 6 }}>¿Qué categoría es?</h2>
          <p style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 16 }}>
            No estoy seguro de la categoría para «{desc || "este movimiento"}». Elígela y la app aprenderá para la próxima.
          </p>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 9 }}>
            {cats.map((c) => (
              <button key={c.id} type="button"
                onClick={() => { finalize(c.id, true); }}
                style={{ padding: "13px 12px", textAlign: "left", cursor: "pointer", fontSize: 14, fontWeight: 600,
                  display: "flex", alignItems: "center", gap: 8, fontFamily: "inherit",
                  background: "var(--paper)", border: "1px solid var(--line)", borderRadius: 14,
                  boxShadow: "var(--shadow-xs)", WebkitTapHighlightColor: "transparent",
                  position: "relative", zIndex: 1 }}>
                <span className="cc-emoji" style={{ fontSize: 19 }}>{c.emoji}</span>{c.name}
              </button>
            ))}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{editing ? "Editar transacción" : "Nueva transacción"}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>

        <div className="cc-tabs" style={{ marginBottom: 4 }}>
          {[["expense", "Gasto"], ["income", "Ingreso"]].map(([k, l]) => (
            <button key={k} className={`cc-tab ${type === k ? "on" : ""}`}
              onClick={() => { setType(k); setCatId("auto"); }}>{l}</button>
          ))}
        </div>

        {/* Sugerencias de movimientos frecuentes */}
        {!editing && (() => {
          const patterns = detectFrequentPatterns(txs, config).filter(p => p.type === type).slice(0, 5);
          if (!patterns.length) return null;
          return (
            <div style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 11, color: "var(--ink-faint)", fontWeight: 600, letterSpacing: ".03em", marginBottom: 6 }}>
                FRECUENTES
              </div>
              <div className="cc-scroll-x" style={{ gap: 6 }}>
                {patterns.map((p, i) => (
                  <button key={i} onClick={() => {
                    setDesc(p.description);
                    setAmount(String(p.amount));
                    if (p.accountId) setAccountId(p.accountId);
                    if (p.categoryId) setCatId(p.categoryId);
                  }}
                    style={{ flexShrink: 0, padding: "8px 14px", borderRadius: 20,
                      background: p.annual ? "rgba(99,102,241,.08)" : "var(--surface)",
                      border: `1px solid ${p.annual ? "rgba(99,102,241,.25)" : "var(--glass-border)"}`,
                      cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 500,
                      color: "var(--ink)", whiteSpace: "nowrap",
                      display: "flex", alignItems: "center", gap: 6 }}>
                    {p.categoryEmoji && <span>{p.categoryEmoji}</span>}
                    <span>{p.description}</span>
                    <span style={{ color: "var(--ink-soft)", fontWeight: 400 }}>{fmtBare(p.amount)}</span>
                    {p.annual && <span style={{ fontSize: 10, color: "#6366F1", fontWeight: 600 }}>ANUAL</span>}
                  </button>
                ))}
              </div>
            </div>
          );
        })()}

        <div className="cc-amount-display">
          <span className="cc-amount-currency">$</span>
          <input className="cc-num" type="text" inputMode="decimal" placeholder="0.00"
            value={amount}
            onChange={(e) => {
              const v = e.target.value.replace(/[^0-9.]/g, "");
              if ((v.match(/\./g) || []).length <= 1) setAmount(v);
            }}
            style={{ width: `${Math.max((amount || "0.00").length, 4)}ch` }} />
          <span className="cc-amount-mxn">mxn</span>
        </div>

        <div style={{ marginBottom: 14 }}>
          <label className="cc-label">Concepto</label>
          <input className="cc-input" placeholder="Ej. tacos con la familia, gasolina, pago de luz…"
            value={desc} onChange={(e) => setDesc(e.target.value)} />
        </div>

        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <div style={{ flex: 1 }}>
            <label className="cc-label">Fecha</label>
            <input className="cc-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
          </div>
          <div style={{ flex: 1 }}>
            <label className="cc-label">Categoría</label>
            <select className="cc-select" value={catId} onChange={(e) => setCatId(e.target.value)}
              disabled={!accountId}>
              <option value="auto">✨ Detectar automáticamente</option>
              {cats.map((c) => <option key={c.id} value={c.id}>{c.emoji} {c.name}</option>)}
            </select>
          </div>
        </div>

        {config.accounts.length > 1 && (
          <div style={{ marginBottom: 14 }}>
            <label className="cc-label">Cuenta</label>
            <select className="cc-select" value={accountId}
              onChange={(e) => { setAccountId(e.target.value); setCatId("auto"); }}
              style={{ borderColor: !accountId ? "var(--gold)" : "var(--line)" }}>
              <option value="">Elegir cuenta…</option>
              {config.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </div>
        )}

        <div style={{ marginBottom: 20 }}>
          {!accountId ? (
            <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>
              Elige primero una cuenta para ver sus categorías.
            </div>
          ) : catId === "auto" && (
            <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>
              La app intentará detectar la categoría con el concepto. Si no está segura, te pregunta.
            </div>
          )}
        </div>

        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 14, fontSize: 15 }}
          disabled={phase === "detecting" || !amount || parseFloat(amount) <= 0 || !accountId}
          onClick={handleSave}>
          {phase === "detecting"
            ? "Detectando categoría…"
            : editing ? "Guardar cambios" : `Guardar ${type === "income" ? "ingreso" : "gasto"}`}
        </button>
      </div>
    </div>
  );
}

/* ===================== ASISTENTE DE CHAT ================================= */
// El historial del chat persiste en memoria mientras la app esté abierta,
// para que no se pierda al cerrar y reabrir el asistente.
let CHAT_MSGS_STORE = null;
let CHAT_HISTORY_STORE = [];
function Assistant({ config, txs, saveConfig, saveTxs, onClose, onOpenImport, autoVoice }) {
  // Detectar patrones para saludo proactivo
  const patterns = detectFrequentPatterns(txs, config);
  const nonRecurring = patterns.filter(p =>
    !(config.recurring || []).find(r =>
      r.description.toLowerCase().trim() === p.description.toLowerCase().trim()));

  const GREET = nonRecurring.length > 0
    ? `¡Hola! Soy tu asistente. 💡 Noté que registras frecuentemente:\n\n${nonRecurring.slice(0, 3).map(p =>
        `• "${p.description}" — ${fmtBare(p.amount)} (${p.count} veces${p.annual ? ", también el año pasado" : ""})`).join("\n")}\n\n¿Quieres que los haga recurrentes para que se registren solos?`
    : "¡Hola! Soy tu asistente. Dime qué quieres y yo lo hago en la app — crear o quitar categorías, cuentas, registrar movimientos… o pregúntame sobre tus gastos.";

  const QUICK = nonRecurring.length > 0
    ? [
        `Hazme recurrente "${nonRecurring[0]?.description}"`,
        "¿Cuánto llevo gastado este mes?",
        "Registra un gasto de 250 en gasolina",
      ]
    : [
        "Crea la categoría Mascotas 🐶",
        "Registra un gasto de 250 en gasolina",
        "¿Cuánto llevo gastado este mes?",
      ];
  const [msgs, setMsgs] = useState(() => CHAT_MSGS_STORE || [{ role: "bot", text: GREET }]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const history = useRef(CHAT_HISTORY_STORE);
  const scroller = useRef(null);

  // mantener el store sincronizado para que persista al cerrar/reabrir
  useEffect(() => { CHAT_MSGS_STORE = msgs; }, [msgs]);

  // ===== Dictado por voz =====
  const [listening, setListening] = useState(false);
  const listeningRef = useRef(false);
  const recRef = useRef(null);
  const finalRef = useRef("");
  const lastHeardRef = useRef("");
  const sendOnEndRef = useRef(false);

  const setListen = (v) => { listeningRef.current = v; setListening(v); };

  const startVoice = () => {
    if (!voiceSupported || busy || listeningRef.current) return;
    try {
      const rec = new SpeechRec();
      rec.lang = "es-MX";
      rec.continuous = true;
      rec.interimResults = true;
      finalRef.current = "";
      lastHeardRef.current = "";
      rec.onresult = (e) => {
        let interim = "";
        let final = "";
        for (let i = e.resultIndex; i < e.results.length; i++) {
          const t = e.results[i][0].transcript;
          if (e.results[i].isFinal) final += t;
          else interim += t;
        }
        if (final) finalRef.current += final;
        const full = (finalRef.current + interim).trim();
        lastHeardRef.current = full; // captura también lo provisional
        setInput(full);
      };
      rec.onerror = () => { setListen(false); };
      rec.onend = () => {
        setListen(false);
        // al terminar (tras soltar) ya están todos los resultados: enviar
        if (sendOnEndRef.current) {
          sendOnEndRef.current = false;
          const text = (lastHeardRef.current || finalRef.current).trim();
          if (text) send(text);
        }
      };
      recRef.current = rec;
      rec.start();
      setListen(true);
      if (navigator.vibrate) navigator.vibrate(8);
    } catch (err) { setListen(false); }
  };

  const stopVoice = (autoSend = true) => {
    if (!listeningRef.current) return;
    sendOnEndRef.current = autoSend; // el envío ocurre en onend (cuando ya hay texto final)
    try { recRef.current && recRef.current.stop(); } catch (e) {}
    if (navigator.vibrate) navigator.vibrate(8);
  };

  // Si se abrió manteniendo presionado el orbe, empieza a escuchar y
  // termina cuando el usuario suelte (en cualquier parte de la pantalla).
  useEffect(() => {
    if (!autoVoice || !voiceSupported) return;
    const t = setTimeout(() => startVoice(), 250);
    const release = () => stopVoice(true);
    window.addEventListener("mouseup", release);
    window.addEventListener("touchend", release);
    return () => {
      clearTimeout(t);
      window.removeEventListener("mouseup", release);
      window.removeEventListener("touchend", release);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (scroller.current) scroller.current.scrollTop = scroller.current.scrollHeight;
  }, [msgs, busy]);

  async function send(text) {
    const userText = (text ?? input).trim();
    if (!userText || busy) return;
    setMsgs((m) => [...m, { role: "me", text: userText }]);
    setInput("");
    setBusy(true);
    history.current.push({ role: "user", content: userText });
    try {
      const raw = await callClaude(assistantSystem(config, txs), history.current);
      const parsed = parseJSON(raw);
      history.current.push({ role: "assistant", content: parsed.message || raw });
      let log = [], charts = [];
      if (Array.isArray(parsed.actions) && parsed.actions.length) {
        // separar show_chart de las demás acciones
        const chartSpecs = parsed.actions.filter((a) => a && a.type === "show_chart");
        const otherActions = parsed.actions.filter((a) => a && a.type !== "show_chart");

        if (otherActions.length) {
          const res = applyActions(config, txs, otherActions);
          saveConfig(res.config);
          saveTxs(res.txs);
          log = res.log;
        }
        for (const cs of chartSpecs) {
          const built = buildChart(cs.spec || cs, config, txs);
          if (built) charts.push(built);
        }
      }
      setMsgs((m) => [...m, { role: "bot", text: parsed.message || "Listo.", log, charts }]);
    } catch (e) {
      setMsgs((m) => [
        ...m,
        { role: "bot", text: "Tuve un problema para procesar eso. ¿Me lo dices de otra forma?" },
      ]);
    }
    setBusy(false);
  }

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
          <h2 className="cc-serif" style={{ fontSize: 21, fontWeight: 600, display: "flex", alignItems: "center", gap: 9 }}>
            <span style={{ width: 10, height: 10, borderRadius: "50%", background: "var(--green)", boxShadow: "0 0 0 4px var(--green-soft)" }} />
            Asistente
          </h2>
          <div style={{ display: "flex", gap: 7 }}>
            {msgs.length > 1 && (
              <button className="cc-btn" style={{ padding: "6px 10px", fontSize: 13 }}
                onClick={() => {
                  setMsgs([{ role: "bot", text: GREET }]);
                  history.current.length = 0;
                  CHAT_MSGS_STORE = null;
                }}
                title="Limpiar conversación">🗑</button>
            )}
            <button className="cc-btn" style={{ padding: "6px 12px", fontSize: 13 }} onClick={onClose}>{t("close")}</button>
          </div>
        </div>

        <div ref={scroller} style={{ display: "flex", flexDirection: "column", gap: 10, overflowY: "auto", maxHeight: "52vh", padding: "2px 2px 6px" }}>
          {msgs.map((m, i) => (
            <div key={i} style={{ display: "flex", flexDirection: "column", gap: 8, alignItems: m.role === "me" ? "flex-end" : "stretch" }}>
              <div className={`cc-bubble ${m.role}`} style={{ whiteSpace: "pre-wrap" }}>
                {m.text}
                {m.log && m.log.length > 0 && (
                  <div style={{ marginTop: 9, display: "flex", flexDirection: "column", gap: 5 }}>
                    {m.log.map((l, j) => (
                      <div key={j} style={{
                        fontSize: 12.5, fontWeight: 600, padding: "5px 9px", borderRadius: 8,
                        background: l.ok ? "rgba(26,122,110,.14)" : "rgba(181,69,58,.12)",
                        color: l.ok ? "var(--green)" : "var(--coral)",
                      }}>
                        {l.ok ? "✓" : "!"} {l.text}
                      </div>
                    ))}
                  </div>
                )}
              </div>
              {/* gráficas fuera de la burbuja, ancho completo */}
              {m.charts && m.charts.length > 0 && m.charts.map((ch, j) => (
                <MiniChart key={j} data={ch} />
              ))}
            </div>
          ))}
          {busy && (
            <div className="cc-bubble bot"><span className="cc-dots"><span /><span /><span /></span></div>
          )}
        </div>

        {msgs.length === 1 && !busy && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 7, margin: "12px 0 2px" }}>
            {QUICK.map((q, i) => (
              <button key={i} className="cc-chip" onClick={() => send(q)}>{q}</button>
            ))}
          </div>
        )}

        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
          <button className="cc-btn" title="Importar desde screenshot"
            onClick={() => { onClose(); onOpenImport && onOpenImport(); }}
            style={{ padding: "10px 12px", fontSize: 18, lineHeight: 1 }}>📸</button>
          {voiceSupported && (
            <button
              className={`cc-btn cc-mic ${listening ? "rec" : ""}`}
              title="Mantén presionado para hablar"
              disabled={busy}
              onMouseDown={(e) => { e.preventDefault(); startVoice(); }}
              onMouseUp={() => stopVoice(true)}
              onMouseLeave={() => listening && stopVoice(true)}
              onTouchStart={(e) => { e.preventDefault(); startVoice(); }}
              onTouchEnd={(e) => { e.preventDefault(); stopVoice(true); }}
              style={{ padding: "10px 13px", fontSize: 18, lineHeight: 1, userSelect: "none", touchAction: "none" }}>
              {listening ? "●" : "🎙"}
            </button>
          )}
          <input
            className="cc-input"
            placeholder={listening ? "Escuchando…" : "Dile algo… ej. quita la categoría Ropa"}
            value={input}
            disabled={busy}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && send()}
          />
          <button className="cc-btn cc-btn-green" disabled={busy || !input.trim()} onClick={() => send()}>
            Enviar
          </button>
        </div>
        {listening && (
          <div style={{ fontSize: 11.5, color: "var(--ink-soft)", textAlign: "center", marginTop: 8 }}>
            Suelta para enviar
          </div>
        )}
      </div>
    </div>
  );
}

/* ===================== MODAL: MOVIMIENTOS RECURRENTES =================== */
function RecurringModal({ config, onClose, onSave }) {
  const rules = config.recurring || [];
  const [view, setView] = useState(rules.length ? "list" : "form"); // list | form
  const [editingId, setEditingId] = useState(null);

  // form state
  const [type, setType] = useState("expense");
  const [amount, setAmount] = useState("");
  const [desc, setDesc] = useState("");
  const [accountId, setAccountId] = useState(config.accounts.length === 1 ? config.accounts[0].id : "");
  const [catId, setCatId] = useState("auto");
  const [freq, setFreq] = useState("monthly");
  const [startDate, setStartDate] = useState(today());
  const [detecting, setDetecting] = useState(false);

  const cats = config.categories.filter((c) => c.type === type && c.accountId === accountId);

  const resetForm = () => {
    setType("expense"); setAmount(""); setDesc("");
    setAccountId(config.accounts.length === 1 ? config.accounts[0].id : "");
    setCatId("auto"); setFreq("monthly"); setStartDate(today()); setEditingId(null);
  };

  const startNew = () => { resetForm(); setView("form"); };

  const startEdit = (r) => {
    setType(r.type); setAmount(String(r.amount)); setDesc(r.description);
    setAccountId(r.accountId); setCatId(r.categoryId || "auto");
    setFreq(r.freq); setStartDate(r.startDate); setEditingId(r.id);
    setView("form");
  };

  const canSave = amount && parseFloat(amount) > 0 && desc.trim() && accountId;

  const save = async () => {
    if (!canSave) return;
    // detección automática de categoría si está en "auto"
    let finalCat = catId === "auto" ? null : (catId || null);
    if (catId === "auto") {
      setDetecting(true);
      try {
        const res = await autoCategorize(desc.trim(), type, accountId, config);
        finalCat = res.catId;
      } catch (e) { finalCat = null; }
      setDetecting(false);
    }
    const rule = {
      id: editingId || uid(),
      type, amount: Math.abs(parseFloat(amount)),
      description: desc.trim(), accountId,
      categoryId: finalCat,
      freq, startDate,
      lastRun: editingId ? (rules.find((r) => r.id === editingId)?.lastRun ?? null) : null,
      active: true,
      paused: false,
    };
    const next = editingId
      ? rules.map((r) => (r.id === editingId ? rule : r))
      : [...rules, rule];
    onSave(next);
    onClose();
  };

  const toggleActive = (id) => {
    onSave(rules.map((r) => {
      if (r.id !== id) return r;
      const nowActive = !isRecActive(r);
      return { ...r, active: nowActive, paused: !nowActive };
    }));
  };

  const remove = (id) => {
    onSave(rules.filter((r) => r.id !== id));
  };

  const accName = (id) => config.accounts.find((a) => a.id === id)?.name || "—";
  const catEmoji = (id) => config.categories.find((c) => c.id === id)?.emoji || "🔁";

  // ===== Vista LISTA =====
  if (view === "list") {
    return (
      <div className="cc-overlay" onClick={onClose}>
        <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
          <div className="cc-grip" />
          <div className="cc-sheet-top">
            <h2>{t("recurringMovements")}</h2>
            <button className="cc-sheet-close" onClick={onClose}>×</button>
          </div>
          <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 16 }}>
            Se generan automáticamente en su fecha. Puedes editarlos como cualquier movimiento.
          </p>

          {rules.length === 0 ? (
            <div style={{ color: "var(--ink-soft)", fontSize: 14, padding: "8px 0 18px" }}>
              Aún no tienes movimientos recurrentes.
            </div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 16 }}>
              {rules.map((r) => (
                <div key={r.id} className="cc-card" style={{ padding: "12px 14px", opacity: isRecActive(r) ? 1 : 0.5 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <div className="cc-emoji" style={{ width: 36, height: 36, borderRadius: 10, background: "var(--surface)",
                      display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>
                      {catEmoji(r.categoryId)}
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: 14, color: "var(--ink)", letterSpacing: "-.01em" }}>{r.description}</div>
                      <div style={{ fontSize: 11.5, color: "var(--ink-soft)", marginTop: 2 }}>
                        {FREQ_LABELS_FN()[r.freq]} · {accName(r.accountId)}
                      </div>
                    </div>
                    <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 14,
                      color: r.type === "income" ? "var(--green)" : "var(--coral)", whiteSpace: "nowrap" }}>
                      {r.type === "income" ? "+" : "−"}{fmtBare(r.amount)}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 7, marginTop: 10 }}>
                    <button className="cc-btn" style={{ flex: 1, padding: "6px 10px", fontSize: 12 }}
                      onClick={() => startEdit(r)}>Editar</button>
                    <button className="cc-btn" style={{ flex: 1, padding: "6px 10px", fontSize: 12 }}
                      onClick={() => toggleActive(r.id)}>{isRecActive(r) ? "Pausar" : "Activar"}</button>
                    <button className="cc-btn" style={{ padding: "6px 10px", fontSize: 12, color: "var(--coral)" }}
                      onClick={() => remove(r.id)}>Eliminar</button>
                  </div>
                </div>
              ))}
            </div>
          )}

          <button className="cc-btn cc-btn-primary" style={{ width: "100%", padding: 13 }}
            onClick={startNew}>＋ Nuevo recurrente</button>
        </div>
      </div>
    );
  }

  // ===== Vista FORM =====
  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{editingId ? "Editar recurrente" : "Nuevo recurrente"}</h2>
          <button className="cc-sheet-close" onClick={() => (rules.length ? setView("list") : onClose())}>×</button>
        </div>

        <div className="cc-tabs" style={{ marginBottom: 4 }}>
          {[["expense", "Gasto"], ["income", "Ingreso"]].map(([k, l]) => (
            <button key={k} className={`cc-tab ${type === k ? "on" : ""}`}
              onClick={() => { setType(k); setCatId(""); }}>{l}</button>
          ))}
        </div>

        <div className="cc-amount-display">
          <span className="cc-amount-currency">$</span>
          <input className="cc-num" type="text" inputMode="decimal" placeholder="0.00"
            value={amount}
            onChange={(e) => {
              const v = e.target.value.replace(/[^0-9.]/g, "");
              if ((v.match(/\./g) || []).length <= 1) setAmount(v);
            }}
            style={{ width: `${Math.max((amount || "0.00").length, 4)}ch` }} />
          <span className="cc-amount-mxn">mxn</span>
        </div>

        <div style={{ marginBottom: 14 }}>
          <label className="cc-label">Concepto</label>
          <input className="cc-input" placeholder="Ej. renta, Netflix, sueldo…"
            value={desc} onChange={(e) => setDesc(e.target.value)} />
        </div>

        <div style={{ marginBottom: 14 }}>
          <label className="cc-label">¿Cada cuánto?</label>
          <div style={{ display: "flex", gap: 7, flexWrap: "wrap" }}>
            {Object.entries(FREQ_LABELS_FN()).map(([k, l]) => (
              <button key={k} onClick={() => setFreq(k)}
                style={{ padding: "9px 14px", borderRadius: 12, cursor: "pointer", fontFamily: "inherit",
                  fontSize: 13, fontWeight: freq === k ? 700 : 500,
                  background: freq === k ? "var(--ink)" : "var(--surface)",
                  color: freq === k ? "#fff" : "var(--ink)",
                  border: `1px solid ${freq === k ? "var(--ink)" : "var(--line)"}` }}>
                {l}
              </button>
            ))}
          </div>
        </div>

        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <div style={{ flex: 1 }}>
            <label className="cc-label">Empieza el</label>
            <input className="cc-input" type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
          </div>
          {config.accounts.length > 1 && (
            <div style={{ flex: 1 }}>
              <label className="cc-label">Cuenta</label>
              <select className="cc-select" value={accountId} onChange={(e) => { setAccountId(e.target.value); setCatId(""); }}>
                <option value="">Elegir…</option>
                {config.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
            </div>
          )}
        </div>

        <div style={{ marginBottom: 18 }}>
          <label className="cc-label">Categoría {accountId ? "" : "(elige cuenta primero)"}</label>
          <select className="cc-select" value={catId} onChange={(e) => setCatId(e.target.value)} disabled={!accountId}>
            <option value="auto">✨ Detectar automáticamente</option>
            <option value="">Sin categoría</option>
            {cats.map((c) => <option key={c.id} value={c.id}>{c.emoji} {c.name}</option>)}
          </select>
          {catId === "auto" && (
            <div style={{ fontSize: 11.5, color: "var(--ink-soft)", marginTop: 6 }}>
              Zafi elegirá la categoría según el concepto al guardar.
            </div>
          )}
        </div>

        <div style={{ fontSize: 12, color: "var(--ink-soft)", marginBottom: 14, lineHeight: 1.5 }}>
          Se generará un {type === "income" ? "ingreso" : "gasto"} de <b>{amount ? fmtBare(parseFloat(amount) || 0) : "$0"}</b> {FREQ_LABELS_FN()[freq].toLowerCase()} a partir del {startDate}. Si la fecha de inicio ya pasó, se crearán los movimientos faltantes al guardar.
        </div>

        <button className="cc-btn cc-btn-primary" style={{ width: "100%", padding: 14, opacity: (canSave && !detecting) ? 1 : 0.4 }}
          disabled={!canSave || detecting} onClick={save}>
          {detecting ? "Detectando categoría…" : (editingId ? "Guardar cambios" : "Crear recurrente")}
        </button>
      </div>
    </div>
  );
}

/* ===================== MODAL: IMPORTAR DESDE SCREENSHOTS ================= */
function ImportModal({ config, txs, onClose, onSave }) {
  const [files, setFiles] = useState([]); // {file, previewUrl}
  const [defaultAccountId, setDefaultAccountId] = useState(
    config.accounts.length === 1 ? config.accounts[0].id : ""
  );
  const [phase, setPhase] = useState("upload"); // upload | processing | review
  const [drafts, setDrafts] = useState([]); // movimientos extraídos editables
  const [error, setError] = useState(null);
  const inputRef = useRef(null);

  const addFiles = (list) => {
    const imgs = Array.from(list || []).filter((f) => f.type.startsWith("image/"));
    const mapped = imgs.map((f) => ({ file: f, previewUrl: URL.createObjectURL(f) }));
    setFiles((prev) => [...prev, ...mapped]);
  };

  const removeFile = (idx) => {
    setFiles((prev) => {
      URL.revokeObjectURL(prev[idx].previewUrl);
      return prev.filter((_, i) => i !== idx);
    });
  };

  async function process() {
    if (!files.length || !defaultAccountId) return;
    setError(null);
    setPhase("processing");
    try {
      const imagesB64 = await Promise.all(files.map((f) => fileToB64(f.file)));
      const acc = config.accounts.find((a) => a.id === defaultAccountId);
      const cats = config.categories
        .filter((c) => c.accountId === defaultAccountId)
        .map((c) => ({ name: c.name, type: c.type, emoji: c.emoji }));

      const sys = `Eres un extractor de movimientos bancarios mexicanos. Vas a recibir una o varias imágenes (screenshots de estados de cuenta o apps bancarias) y debes extraer TODOS los movimientos visibles.

Cuenta destino: ${acc ? acc.name : "Principal"}
Categorías disponibles para esa cuenta:
${JSON.stringify(cats)}

Responde SIEMPRE con UN SOLO objeto JSON sin markdown:
{"movements":[{"date":"YYYY-MM-DD","amount":<número positivo>,"type":"income"|"expense","description":"comercio o concepto","categoryName":"<nombre de categoría existente, o null si no estás seguro>"},...]}

REGLAS GENERALES:
- "date" en formato YYYY-MM-DD. Si la imagen solo muestra día/mes, usa el año actual (${today().slice(0,4)}). Si no hay fecha, usa la de hoy: ${today()}.
- "amount" SIEMPRE positivo, sin signo.
- "type": "expense" si es un cargo/gasto/débito; "income" si es un abono/depósito/crédito.
- "description": el comercio o concepto que aparece, breve (ej. "OXXO", "Uber", "Pago tarjeta", "Spotify").
- "categoryName": elige una de las categorías de arriba que coincida con el tipo (gasto/ingreso). Si no estás seguro, usa null.
- Ignora encabezados, saldos totales, comisiones que no sean movimientos reales.
- Si la imagen no muestra movimientos, devuelve {"movements":[]}.

🚨 CÓMO ASIGNAR LA FECHA — REGLA CRÍTICA, LEE CON CUIDADO:
Muchas apps bancarias mexicanas (Santander, BBVA, Banorte) muestran los separadores de fecha como un ENCABEZADO que aplica a los movimientos que aparecen DEBAJO de ese separador, NO a los de arriba.

Ejemplo del formato típico de Santander:
  [movimiento A] ←  ¿de qué día es?
  ─── sábado 02 de mayo 2026 ───
  [movimiento B]
  [movimiento C]
  ─── viernes 01 de mayo 2026 ───
  [movimiento D]

En este ejemplo:
- B y C son del SÁBADO 02 de mayo (debajo de su separador).
- D es del VIERNES 01 de mayo.
- A NO se le puede asignar fecha con certeza desde esta sola imagen porque su separador está cortado/fuera de pantalla (probablemente sea de un día más reciente, como lunes 04).

REGLAS DE FECHA:
1. La fecha de cada movimiento es la del separador que aparece JUSTO ARRIBA de él, o el más cercano hacia arriba.
2. Si un movimiento aparece en la parte SUPERIOR de la imagen sin un separador de fecha visible arriba de él (porque la captura empieza en medio del scroll), NO INVENTES su fecha. En ese caso, omite ese movimiento — aparecerá completo en otro screenshot del lote.
3. Las fechas dentro de la app van de más recientes (arriba) a más antiguas (abajo). Si ves un separador "sábado 02 mayo" y arriba hay movimientos sin separador visible, esos movimientos son de un día POSTERIOR al 02 (lunes 04, martes 05, etc.), no anterior.

🚨 EVITA DUPLICADOS ENTRE IMÁGENES:
Cuando subo varios screenshots de la misma app, los movimientos se traslapan entre fotos. Si detectas el mismo movimiento (mismo monto, mismo comercio) en dos imágenes distintas, inclúyelo UNA SOLA VEZ con la fecha correcta de su separador. Combina la información de ambas imágenes: si en una imagen el movimiento aparece cortado en la parte superior (sin fecha visible) y en otra imagen aparece bajo un separador claro, usa la fecha del separador.

- Incluye TODOS los movimientos que puedas datar con certeza, en el orden en que aparecen.`;

      const raw = await callClaudeVision(sys, "Extrae todos los movimientos visibles.", imagesB64);
      const parsed = parseJSON(raw);
      const list = Array.isArray(parsed.movements) ? parsed.movements : [];
      if (!list.length) {
        setError("No detecté movimientos en las imágenes. Verifica que se vean claros.");
        setPhase("upload");
        return;
      }
      // mapear a drafts con resolución de categoría y detección de duplicados
      const ds = list.map((m, i) => {
        const c = m.categoryName ? findCat({ ...config }, m.categoryName, defaultAccountId) : null;
        const candidate = {
          date: m.date || today(),
          amount: Math.abs(Number(m.amount) || 0),
          type: m.type === "income" ? "income" : "expense",
          description: m.description || "",
          accountId: defaultAccountId,
        };
        const dup = findDuplicate(candidate, txs || []);
        return {
          tempId: "draft" + i,
          ...candidate,
          categoryId: c ? c.id : "",
          selected: !dup, // desmarcar duplicados por defecto
          duplicate: dup ? { date: dup.date, amount: dup.amount, description: dup.description } : null,
        };
      });
      setDrafts(ds);
      setPhase("review");
    } catch (e) {
      console.error(e);
      setError("Hubo un problema procesando las imágenes. Inténtalo de nuevo.");
      setPhase("upload");
    }
  }

  function updateDraft(tempId, patch) {
    setDrafts((prev) => prev.map((d) => (d.tempId === tempId ? { ...d, ...patch } : d)));
  }

  function saveAll() {
    const valid = drafts.filter((d) => d.selected && d.amount > 0);
    if (!valid.length) return;
    // aprender de cada categoría asignada manualmente
    let cats = config.categories;
    valid.forEach((d) => {
      if (!d.categoryId || !d.description) return;
      const target = cats.find((c) => c.id === d.categoryId);
      if (!target) return;
      const kws = extractKW(d.description).slice(0, 3);
      if (!kws.length) return;
      const sameName = norm(target.name).trim();
      cats = cats.map((c) =>
        norm(c.name).trim() === sameName && c.type === target.type
          ? { ...c, keywords: [...new Set([...(c.keywords || []), ...kws])] }
          : c
      );
    });
    const txs = valid.map((d) => ({
      id: uid(), type: d.type, amount: d.amount, description: d.description,
      categoryId: d.categoryId || null, accountId: d.accountId, date: d.date,
    }));
    onSave(txs, cats);
  }

  /* ---------- pantalla: subir ---------- */
  if (phase === "upload") {
    return (
      <div className="cc-overlay" onClick={onClose}>
        <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
          <div className="cc-grip" />
          <div className="cc-sheet-top">
            <h2>Importar screenshot</h2>
            <button className="cc-sheet-close" onClick={onClose}>×</button>
          </div>
          <p style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 16 }}>
            Sube screenshots de tu app bancaria o estado de cuenta. La IA extrae los movimientos y tú los revisas.
          </p>

          {config.accounts.length > 1 && (
            <div style={{ marginBottom: 14 }}>
              <label className="cc-label">Cuenta destino</label>
              <select className="cc-select" value={defaultAccountId} onChange={(e) => setDefaultAccountId(e.target.value)}
                style={{ borderColor: !defaultAccountId ? "var(--gold)" : "var(--line)" }}>
                <option value="">Elegir cuenta…</option>
                {config.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
            </div>
          )}

          <input ref={inputRef} type="file" accept="image/*" multiple style={{ display: "none" }}
            onChange={(e) => addFiles(e.target.files)} />

          <button className="cc-card" onClick={() => inputRef.current?.click()}
            style={{ width: "100%", padding: 26, textAlign: "center", border: "2px dashed var(--line)",
              background: "var(--surface-2)", cursor: "pointer", marginBottom: 14 }}>
            <div style={{ fontSize: 30, marginBottom: 6 }}>📷</div>
            <div style={{ fontWeight: 600, fontSize: 14 }}>Toca para elegir imágenes</div>
            <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 4 }}>
              También puedes seleccionar varias a la vez
            </div>
          </button>

          {files.length > 0 && (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, marginBottom: 16 }}>
              {files.map((f, i) => (
                <div key={i} style={{ position: "relative", borderRadius: 10, overflow: "hidden", aspectRatio: "1", border: "1px solid var(--line)" }}>
                  <img src={f.previewUrl} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
                  <button onClick={() => removeFile(i)}
                    style={{ position: "absolute", top: 4, right: 4, width: 24, height: 24, borderRadius: "50%",
                      border: "none", background: "rgba(0,0,0,.7)", color: "#fff", cursor: "pointer", fontSize: 14, lineHeight: 1 }}>✕</button>
                </div>
              ))}
            </div>
          )}

          {error && (
            <div style={{ padding: 11, borderRadius: 10, background: "var(--coral-soft)", color: "var(--coral)", fontSize: 13, marginBottom: 14, fontWeight: 600 }}>
              {error}
            </div>
          )}

          <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 14, fontSize: 15 }}
            disabled={!files.length || !defaultAccountId} onClick={process}>
            ✨ Extraer movimientos
          </button>
        </div>
      </div>
    );
  }

  /* ---------- pantalla: procesando ---------- */
  if (phase === "processing") {
    return (
      <div className="cc-overlay">
        <div className="cc-sheet" style={{ padding: "40px 18px" }}>
          <div className="cc-grip" />
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <div style={{ fontSize: 36, marginBottom: 14 }}>🔍</div>
            <div className="cc-serif" style={{ fontSize: 20, fontWeight: 600, marginBottom: 6 }}>
              Leyendo tus movimientos…
            </div>
            <div style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 18 }}>
              La IA está revisando {files.length} imagen{files.length > 1 ? "es" : ""}
            </div>
            <div className="cc-dots"><span /><span /><span /></div>
          </div>
        </div>
      </div>
    );
  }

  /* ---------- pantalla: revisar ---------- */
  const okCount = drafts.filter((d) => d.selected && d.amount > 0).length;
  const accCats = (type) => config.categories.filter((c) => c.accountId === defaultAccountId && c.type === type);

  return (
    <ReviewScreen
      drafts={drafts}
      updateDraft={updateDraft}
      accCats={accCats}
      onBack={() => { setDrafts([]); setPhase("upload"); }}
      onSave={saveAll}
      onClose={onClose}
      sourceLabel="Encontré"
    />
  );
}

/* ============================ ESTADÍSTICAS =============================== */
/* ===== Gráfica de categorías versátil: pastel / dona / barras ===== */
const CHART_PALETTE = ["#5B6EE8", "#7C8BF5", "#60A5FA", "#5EEAD4", "#A78BFA", "#F0A868", "#E8849B", "#7E8AA0", "#86B98E", "#C9A24B"];

function CategoryChart({ rows, type, onPick }) {
  // rows: [{cat, amt}]
  const [chartType, setChartType] = useState(type === "income" ? "donut" : "bars");
  const [activeIdx, setActiveIdx] = useState(null);
  const total = rows.reduce((s, r) => s + r.amt, 0);
  if (!rows.length) return <div style={{ color: "var(--ink-soft)", fontSize: 13, padding: "8px 0 14px" }}>No hay datos en el periodo.</div>;

  const data = rows.map((r, i) => ({ ...r, color: CHART_PALETTE[i % CHART_PALETTE.length] }));
  const accentColor = type === "income" ? "var(--green)" : "var(--coral)";

  const TYPES = [["bars", "Barras"], ["pie", "Pastel"], ["donut", "Dona"]];

  return (
    <div>
      {/* switch de tipo */}
      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        {TYPES.map(([k, l]) => (
          <button key={k} onClick={() => { setChartType(k); setActiveIdx(null); }}
            style={{ padding: "5px 12px", borderRadius: 10, cursor: "pointer", fontFamily: "inherit",
              fontSize: 11.5, fontWeight: chartType === k ? 700 : 500,
              background: chartType === k ? "var(--ink)" : "var(--surface)",
              color: chartType === k ? "#fff" : "var(--ink-soft)",
              border: `1px solid ${chartType === k ? "var(--ink)" : "var(--line)"}` }}>
            {l}
          </button>
        ))}
      </div>

      {(chartType === "pie" || chartType === "donut") && (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 14 }}>
          <CatPie data={data} total={total} donut={chartType === "donut"}
            active={activeIdx} onHover={setActiveIdx} />
          <div style={{ display: "flex", flexDirection: "column", gap: 2, width: "100%" }}>
            {data.map((d, i) => {
              const pct = total ? Math.round((d.amt / total) * 100) : 0;
              return (
                <button key={d.cat.id} onClick={() => onPick && onPick(d.cat.id)}
                  onMouseEnter={() => setActiveIdx(i)} onMouseLeave={() => setActiveIdx(null)}
                  style={{ display: "flex", alignItems: "center", gap: 9, padding: "8px 6px",
                    background: activeIdx === i ? "var(--surface)" : "transparent", border: "none",
                    borderRadius: 8, cursor: "pointer", fontFamily: "inherit", textAlign: "left", width: "100%",
                    transition: "background .12s" }}>
                  <span style={{ width: 10, height: 10, borderRadius: 3, background: d.color, flexShrink: 0 }} />
                  <span className="cc-emoji" style={{ fontSize: 15 }}>{d.cat.emoji}</span>
                  <span style={{ flex: 1, fontSize: 13, fontWeight: 600, color: "var(--ink)" }}>{d.cat.name}</span>
                  <span className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 13, color: accentColor }}>
                    {fmtBare(d.amt)}
                  </span>
                  <span style={{ fontSize: 10.5, color: "var(--ink-faint)", width: 30, textAlign: "right" }}>{pct}%</span>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {chartType === "bars" && (
        <div style={{ display: "flex", flexDirection: "column" }}>
          {data.map((d) => {
            const pct = total ? Math.round((d.amt / total) * 100) : 0;
            const maxAmt = data[0].amt;
            const w = maxAmt ? (d.amt / maxAmt) * 100 : 0;
            return (
              <button key={d.cat.id} onClick={() => onPick && onPick(d.cat.id)}
                style={{ display: "flex", alignItems: "center", gap: 10, padding: "11px 0",
                  background: "transparent", border: "none", borderBottom: "1px solid var(--line-soft)",
                  cursor: "pointer", fontFamily: "inherit", textAlign: "left", width: "100%" }}>
                <div style={{ width: 34, height: 34, borderRadius: 10, background: "var(--surface)",
                  display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>
                  <span className="cc-emoji">{d.cat.emoji}</span>
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 600, fontSize: 13, color: "var(--ink)", letterSpacing: "-.01em" }}>{d.cat.name}</div>
                  <div style={{ height: 5, background: "var(--surface-2)", borderRadius: 99, overflow: "hidden", marginTop: 5 }}>
                    <div style={{ height: "100%", width: `${w}%`, background: type === "income" ? "var(--green)" : "var(--bar-fill)", borderRadius: 99 }} />
                  </div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 14, color: accentColor }}>
                    {fmtBare(d.amt)}<span style={{ fontSize: 10, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
                  </div>
                  <div style={{ fontSize: 10.5, color: "var(--ink-faint)", marginTop: 1 }}>{pct}%</div>
                </div>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* Pastel/dona interactivo con segmento activo resaltado */
function CatPie({ data, total, donut, active, onHover }) {
  const size = 180, cx = size / 2, cy = size / 2, r = 78, ir = donut ? 46 : 0;
  let acc = 0;
  const slices = data.map((d) => {
    const frac = total ? d.amt / total : 0;
    const start = acc; acc += frac;
    return { ...d, start, end: acc };
  });
  function arc(s, e, rad, innerRad) {
    const a0 = s * Math.PI * 2 - Math.PI / 2;
    const a1 = e * Math.PI * 2 - Math.PI / 2;
    const large = e - s > 0.5 ? 1 : 0;
    const x0 = cx + rad * Math.cos(a0), y0 = cy + rad * Math.sin(a0);
    const x1 = cx + rad * Math.cos(a1), y1 = cy + rad * Math.sin(a1);
    if (innerRad === 0) {
      return `M${cx},${cy} L${x0},${y0} A${rad},${rad} 0 ${large} 1 ${x1},${y1} Z`;
    }
    const xi0 = cx + innerRad * Math.cos(a0), yi0 = cy + innerRad * Math.sin(a0);
    const xi1 = cx + innerRad * Math.cos(a1), yi1 = cy + innerRad * Math.sin(a1);
    return `M${x0},${y0} A${rad},${rad} 0 ${large} 1 ${x1},${y1} L${xi1},${yi1} A${innerRad},${innerRad} 0 ${large} 0 ${xi0},${yi0} Z`;
  }
  const activeSlice = active != null ? slices[active] : null;
  return (
    <svg viewBox={`0 0 ${size} ${size}`} style={{ width: 180, height: 180, flexShrink: 0 }}>
      {slices.map((s, i) =>
        s.end - s.start < 0.001 ? null :
        <path key={i} d={arc(s.start, s.end, active === i ? r + 4 : r, ir)} fill={s.color}
          opacity={active == null || active === i ? 1 : 0.45}
          style={{ transition: "opacity .15s, d .15s", cursor: "pointer" }}
          onMouseEnter={() => onHover && onHover(i)} onMouseLeave={() => onHover && onHover(null)} />
      )}
      {donut && (
        <>
          <text x={cx} y={cy - 4} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="12" fill="var(--ink-soft)">
            {activeSlice ? activeSlice.cat.name : "Total"}
          </text>
          <text x={cx} y={cy + 15} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="16" fontWeight="600" fill="var(--ink)">
            {fmtBare(activeSlice ? activeSlice.amt : total)}
          </text>
        </>
      )}
    </svg>
  );
}

function Estadisticas({ config, txs, dateRange, onEdit, saveConfig }) {
  const [view, setView] = useState("all"); // all | accountId
  const [detail, setDetail] = useState(null); // {kind, label, color, txs, total}
  const [configOpen, setConfigOpen] = useState(false);

  // secciones configurables (orden + visibilidad)
  const STATS_DEFAULT = [
    { id: "summary", label: "Resumen ingresos/gastos", on: true },
    { id: "kpis", label: "Indicadores (KPIs)", on: true },
    { id: "incCats", label: "Ingresos por categoría", on: true },
    { id: "expCats", label: "Gastos por categoría", on: true },
    { id: "trend", label: "Evolución de saldo", on: true },
    { id: "topCat", label: "En lo que más gastaste", on: true },
  ];
  const statsSections = (() => {
    const saved = config.statsSections || [];
    // merge: respeta orden guardado, agrega nuevas que falten
    const byId = Object.fromEntries(saved.map((s) => [s.id, s]));
    const merged = saved
      .filter((s) => STATS_DEFAULT.some((d) => d.id === s.id))
      .map((s) => ({ ...STATS_DEFAULT.find((d) => d.id === s.id), ...s }));
    STATS_DEFAULT.forEach((d) => { if (!byId[d.id]) merged.push(d); });
    return merged.length ? merged : STATS_DEFAULT;
  })();
  const saveStatsSections = (next) => saveConfig({ ...config, statsSections: next });

  const scopedTxs = view === "all" ? txs : txs.filter((t) => t.accountId === view);
  const scopedInitial = view === "all"
    ? config.accounts.reduce((s, a) => s + (a.initialBalance || 0), 0)
    : (config.accounts.find((a) => a.id === view)?.initialBalance || 0);

  // ============== datos del rango ==============
  const rangeTxs = txsInRange(scopedTxs, dateRange);
  const rangeStat = statTxs(rangeTxs).all;
  const rangeIncome = rangeStat.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const rangeExpense = rangeStat.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  const rangeFlow = rangeIncome - rangeExpense;

  // categorías de gasto e ingreso por separado
  const expByCat = {};
  const incByCat = {};
  rangeStat.forEach((t) => {
    if (!t.categoryId) return;
    if (t.type === "expense") expByCat[t.categoryId] = (expByCat[t.categoryId] || 0) + t.amount;
    else incByCat[t.categoryId] = (incByCat[t.categoryId] || 0) + t.amount;
  });
  const expRows = Object.entries(expByCat)
    .map(([id, amt]) => ({ cat: config.categories.find((c) => c.id === id), amt }))
    .filter((x) => x.cat).sort((a, b) => b.amt - a.amt);
  const incRows = Object.entries(incByCat)
    .map(([id, amt]) => ({ cat: config.categories.find((c) => c.id === id), amt }))
    .filter((x) => x.cat).sort((a, b) => b.amt - a.amt);

  // pie de gastos (siempre del rango)
  const pieTotal = rangeExpense;

  // ============== evolución de saldo (90 días o todo el rango) ==============
  const { from: rfrom, to: rto } = resolveRange(dateRange);
  const sorted = [...scopedTxs].sort((a, b) => a.date.localeCompare(b.date));
  const balancePoints = [];
  let running = scopedInitial;
  for (const t of sorted) {
    if (t.date < rfrom) running += t.type === "income" ? t.amount : -t.amount;
  }
  const dayMap = new Map();
  dayMap.set(rfrom, running);
  for (const t of sorted) {
    if (t.date < rfrom || t.date > rto) continue;
    running += t.type === "income" ? t.amount : -t.amount;
    dayMap.set(t.date, running);
  }
  let lastVal = scopedInitial + sorted.filter((t) => t.date < rfrom).reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
  const startD = new Date(rfrom + "T12:00:00");
  const endD = new Date(rto + "T12:00:00");
  for (let d = new Date(startD); d <= endD; d.setDate(d.getDate() + 1)) {
    const k = d.toISOString().slice(0, 10);
    if (dayMap.has(k)) lastVal = dayMap.get(k);
    balancePoints.push({ date: k, val: lastVal });
  }

  // ============== KPIs ==============
  const days = Math.max(1, Math.round((endD - startD) / 86400000) + 1);
  const avgDaily = rangeExpense / days;
  const topCat = expRows[0];
  const txCount = rangeTxs.length;

  const showName = view === "all" ? "General" : (config.accounts.find((a) => a.id === view)?.name || "");

  // helpers de drill-down
  function openDetail(kind) {
    if (kind === "income") {
      const list = rangeStat.filter((t) => t.type === "income");
      setDetail({ kind: "income", label: "Todos los ingresos", color: "var(--green)",
        emoji: "📥", txs: list, total: rangeIncome });
    } else if (kind === "expense") {
      const list = rangeStat.filter((t) => t.type === "expense");
      setDetail({ kind: "expense", label: "Todos los gastos", color: "var(--coral)",
        emoji: "📤", txs: list, total: rangeExpense });
    }
  }
  function openCategoryDetail(catId) {
    const c = config.categories.find((x) => x.id === catId);
    if (!c) return;
    const list = rangeStat.filter((t) => t.categoryId === catId);
    const total = list.reduce((s, t) => s + t.amount, 0);
    setDetail({ kind: "category", label: c.name, color: c.type === "income" ? "var(--green)" : "var(--coral)",
      emoji: c.emoji, txs: list, total });
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      {/* selector de cuenta */}
      <div className="cc-fade">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div className="cc-label" style={{ marginBottom: 0 }}>Ver estadísticas de</div>
          <button className="cc-gear" onClick={() => setConfigOpen(true)} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, padding: "6px 12px", width: "auto" }}>
            ⚙️ Personalizar
          </button>
        </div>
        <div className="cc-scroll-x">
          <button className={`cc-acc-card ${view === "all" ? "on" : ""}`} onClick={() => setView("all")}
            style={{ minWidth: 120 }}>
            <div className="cc-acc-label">Todas</div>
            <div className="cc-acc-name">General</div>
          </button>
          {config.accounts.map((a) => (
            <button key={a.id} className={`cc-acc-card ${view === a.id ? "on" : ""}`} onClick={() => setView(a.id)}
              style={{ minWidth: 120 }}>
              <div className="cc-acc-label">🏦 Cuenta</div>
              <div className="cc-acc-name">{a.name}</div>
            </button>
          ))}
        </div>
      </div>

      {rangeTxs.length === 0 ? (
        <div className="cc-card" style={{ padding: 26, textAlign: "center" }}>
          <div style={{ fontSize: 30, marginBottom: 8 }}>📊</div>
          <div style={{ fontWeight: 600, fontSize: 15 }}>Sin movimientos en {showName}</div>
          <div style={{ fontSize: 13, color: "var(--ink-soft)", marginTop: 6 }}>
            Para el periodo: <b>{rangeLabel(dateRange)}</b>. Cambia el rango arriba o registra movimientos.
          </div>
        </div>
      ) : (
        <>
          {statsSections.filter((s) => s.on).map((s) => {
            if (s.id === "summary") return (
              <div key={s.id} className="cc-card" style={{ padding: 18 }}>
                <div className="cc-label" style={{ marginBottom: 12 }}>Resumen · {rangeLabel(dateRange)}</div>
                <div style={{ display: "flex", gap: 10 }}>
                  <button onClick={() => openDetail("income")}
                    style={{ flex: 1, padding: "14px 15px", background: "var(--surface)",
                      border: "1px solid var(--line-soft)", borderRadius: 14, cursor: "pointer", textAlign: "left",
                      fontFamily: "inherit", display: "flex", flexDirection: "column", gap: 4 }}>
                    <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
                      textTransform: "uppercase", letterSpacing: ".06em" }}>Ingresos</div>
                    <div className="cc-serif cc-num" style={{ fontSize: 22, fontWeight: 500, color: "var(--green)" }}>{fmtBare(rangeIncome)}</div>
                    <div style={{ fontSize: 11, color: "var(--ink-faint)" }}>Tocar para detalle ▸</div>
                  </button>
                  <button onClick={() => openDetail("expense")}
                    style={{ flex: 1, padding: "14px 15px", background: "var(--surface)",
                      border: "1px solid var(--line-soft)", borderRadius: 14, cursor: "pointer", textAlign: "left",
                      fontFamily: "inherit", display: "flex", flexDirection: "column", gap: 4 }}>
                    <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
                      textTransform: "uppercase", letterSpacing: ".06em" }}>Gastos</div>
                    <div className="cc-serif cc-num" style={{ fontSize: 22, fontWeight: 500, color: "var(--coral)" }}>{fmtBare(rangeExpense)}</div>
                    <div style={{ fontSize: 11, color: "var(--ink-faint)" }}>Tocar para detalle ▸</div>
                  </button>
                </div>
                <div style={{ marginTop: 12, paddingTop: 12, borderTop: "1px solid var(--line-soft)",
                  display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 13, fontWeight: 600, color: "var(--ink-soft)" }}>Flujo neto</span>
                  <span className="cc-serif cc-num" style={{ fontSize: 18, fontWeight: 500,
                    color: rangeFlow >= 0 ? "var(--ink)" : "var(--coral)" }}>
                    {rangeFlow >= 0 ? "+" : "−"}{fmtBare(Math.abs(rangeFlow))}<span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 11, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
                  </span>
                </div>
              </div>
            );

            if (s.id === "kpis") return (
              <div key={s.id} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <KpiCard label="Gasto promedio diario" value={fmt(avgDaily)} color="var(--coral)" />
                <KpiCard label="Movimientos en el periodo" value={String(txCount)} color="var(--ink)" />
              </div>
            );

            if (s.id === "incCats" && incRows.length > 0) return (
              <div key={s.id} className="cc-card" style={{ padding: 18 }}>
                <div className="cc-label" style={{ marginBottom: 12 }}>Ingresos por categoría</div>
                <CategoryChart rows={incRows} type="income" onPick={openCategoryDetail} />
              </div>
            );

            if (s.id === "expCats" && expRows.length > 0) return (
              <div key={s.id}>
                <div className="cc-card" style={{ padding: 18 }}>
                  <div className="cc-label" style={{ marginBottom: 12 }}>Gastos por categoría</div>
                  <CategoryChart rows={expRows} type="expense" onPick={openCategoryDetail} />
                </div>
                {/* Gráfica de líneas por categoría */}
                <div className="cc-card" style={{ padding: 18, marginTop: 10 }}>
                  <div className="cc-label" style={{ marginBottom: 10 }}>Tendencia por categoría</div>
                  <CategoryTrendChart txs={scopedTxs} dateRange={dateRange} config={config} />
                </div>
              </div>
            );

            if (s.id === "trend" && balancePoints.length >= 2) return (
              <div key={s.id} className="cc-card" style={{ padding: 18 }}>
                <div className="cc-label" style={{ marginBottom: 10 }}>Saldo · {rangeLabel(dateRange)}</div>
                <LineChart points={balancePoints} />
                <div style={{ fontSize: 11, color: "var(--ink-faint)", marginTop: 8, textAlign: "center" }}>
                  Desliza sobre la gráfica para ver cualquier día
                </div>
              </div>
            );

            if (s.id === "topCat" && topCat) return (
              <div key={s.id} className="cc-card" style={{ padding: 18 }}>
                <div className="cc-label" style={{ marginBottom: 6 }}>En lo que más gastaste</div>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <span className="cc-emoji" style={{ fontSize: 32 }}>{topCat.cat.emoji}</span>
                  <div style={{ flex: 1 }}>
                    <div className="cc-serif" style={{ fontSize: 19, fontWeight: 600 }}>{topCat.cat.name}</div>
                    <div className="cc-num" style={{ fontSize: 14, color: "var(--ink-soft)" }}>
                      {fmt(topCat.amt)} · {pieTotal ? Math.round((topCat.amt / pieTotal) * 100) : 0}% de tus gastos
                    </div>
                  </div>
                </div>
              </div>
            );

            return null;
          })}
        </>
      )}

      {configOpen && (
        <StatsConfigModal
          sections={statsSections}
          onClose={() => setConfigOpen(false)}
          onSave={(next) => { saveStatsSections(next); setConfigOpen(false); }}
        />
      )}

      {detail && (
        <DetailModal config={config} detail={detail} dateRange={dateRange}
          onClose={() => setDetail(null)}
          onEditTx={(t) => { setDetail(null); onEdit(t); }} />
      )}
    </div>
  );
}

/* ===== Modal: personalizar secciones de Estadísticas (reordenar + on/off) ===== */
function StatsConfigModal({ sections, onClose, onSave }) {
  const [items, setItems] = useState(sections.map((s) => ({ ...s })));

  const move = (i, dir) => {
    const j = i + dir;
    if (j < 0 || j >= items.length) return;
    const next = [...items];
    [next[i], next[j]] = [next[j], next[i]];
    setItems(next);
  };
  const toggle = (i) => {
    setItems(items.map((s, idx) => (idx === i ? { ...s, on: !s.on } : s)));
  };

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{t("customizeStatsTitle")}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 16 }}>
          Reordena con las flechas y muestra u oculta cada sección.
        </p>

        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 18 }}>
          {items.map((s, i) => (
            <div key={s.id} className="cc-card" style={{ padding: "11px 14px", display: "flex", alignItems: "center", gap: 10, opacity: s.on ? 1 : 0.55 }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                <button onClick={() => move(i, -1)} disabled={i === 0}
                  style={{ background: "var(--surface)", border: "1px solid var(--line)", borderRadius: 6, width: 26, height: 20, cursor: i === 0 ? "default" : "pointer", opacity: i === 0 ? 0.3 : 1, fontSize: 11, lineHeight: 1, color: "var(--ink)" }}>▲</button>
                <button onClick={() => move(i, 1)} disabled={i === items.length - 1}
                  style={{ background: "var(--surface)", border: "1px solid var(--line)", borderRadius: 6, width: 26, height: 20, cursor: i === items.length - 1 ? "default" : "pointer", opacity: i === items.length - 1 ? 0.3 : 1, fontSize: 11, lineHeight: 1, color: "var(--ink)" }}>▼</button>
              </div>
              <span style={{ flex: 1, fontSize: 14, fontWeight: 600, color: "var(--ink)" }}>{s.label}</span>
              <button onClick={() => toggle(i)}
                style={{ width: 46, height: 27, borderRadius: 99, border: "none", cursor: "pointer", position: "relative",
                  background: s.on ? "var(--green)" : "var(--surface-2)", transition: "background .15s" }}>
                <span style={{ position: "absolute", top: 3, left: s.on ? 22 : 3, width: 21, height: 21, borderRadius: "50%",
                  background: "#fff", transition: "left .15s", boxShadow: "0 1px 3px rgba(0,0,0,.2)" }} />
              </button>
            </div>
          ))}
        </div>

        <button className="cc-btn cc-btn-primary" style={{ width: "100%", padding: 14 }}
          onClick={() => onSave(items)}>Guardar</button>
      </div>
    </div>
  );
}

function KpiCard({ label, value, sub, color }) {
  return (
    <div className="cc-card cc-card-boxed" style={{ padding: 14, paddingTop: 14 }}>
      <div className="cc-label" style={{ marginBottom: 6 }}>{label}</div>
      <div className="cc-serif cc-num" style={{ fontSize: 22, fontWeight: 500, color: color || "var(--ink)", letterSpacing: "-.025em", lineHeight: 1 }}>{value}</div>
      {sub && <div className="cc-num" style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

/* ----------------------- gráfica de línea (SVG) -------------------------- */
/* CategoryTrendChart: gráfica de líneas multi-categoría */
const CAT_LINE_COLORS = ["#F87171","#60A5FA","#34D399","#FBBF24","#A78BFA","#FB923C"];
function CategoryTrendChart({ txs, dateRange, config }) {
  const [hover, setHover] = useState(null);
  const rangeTx = txsInRange(txs, dateRange).filter(t => t.type === "expense" && t.categoryId);
  if (rangeTx.length < 2) return <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>Datos insuficientes.</div>;

  // top 5 categorías por gasto total
  const totals = {};
  rangeTx.forEach(t => { totals[t.categoryId] = (totals[t.categoryId] || 0) + t.amount; });
  const topCatIds = Object.entries(totals).sort((a, b) => b[1] - a[1]).slice(0, 5).map(e => e[0]);

  // compute daily cumulative per category
  const dates = [...new Set(rangeTx.map(t => t.date))].sort();
  if (dates.length < 2) return <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>Datos insuficientes.</div>;
  const startD = new Date(dates[0]), endD = new Date(dates[dates.length - 1]);
  const allDays = [];
  for (let d = new Date(startD); d <= endD; d.setDate(d.getDate() + 1)) allDays.push(d.toISOString().slice(0, 10));

  const series = topCatIds.map((catId, ci) => {
    const cat = config.categories.find(c => c.id === catId);
    const daily = {};
    rangeTx.filter(t => t.categoryId === catId).forEach(t => { daily[t.date] = (daily[t.date] || 0) + t.amount; });
    let cum = 0;
    const pts = allDays.map(d => { cum += daily[d] || 0; return { date: d, val: cum }; });
    return { catId, name: cat?.name || "?", emoji: cat?.emoji || "❔", color: CAT_LINE_COLORS[ci], pts };
  });

  const W = 600, H = 200, P = 12, PB = 6;
  const allVals = series.flatMap(s => s.pts.map(p => p.val));
  const maxV = Math.max(...allVals) || 1;
  const xOf = (i) => P + (i / Math.max(allDays.length - 1, 1)) * (W - P * 2);
  const yOf = (v) => H - PB - (v / maxV) * (H - P - PB);

  const handleMove = (e) => {
    const svg = e.currentTarget;
    const rect = svg.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const x = ((clientX - rect.left) / rect.width) * W;
    const i = Math.round(((x - P) / (W - P * 2)) * (allDays.length - 1));
    setHover(Math.max(0, Math.min(allDays.length - 1, i)));
  };

  return (
    <div>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto" }}
        onMouseMove={handleMove} onTouchMove={handleMove} onMouseLeave={() => setHover(null)} onTouchEnd={() => setHover(null)}>
        {/* grid lines */}
        {[0.25, 0.5, 0.75].map(f => (
          <line key={f} x1={P} x2={W - P} y1={yOf(maxV * f)} y2={yOf(maxV * f)}
            stroke="var(--line-soft)" strokeWidth=".5" strokeDasharray="4,4" />
        ))}
        {/* lines per category */}
        {series.map(s => {
          const path = s.pts.map((p, i) => `${i === 0 ? "M" : "L"}${xOf(i).toFixed(1)},${yOf(p.val).toFixed(1)}`).join(" ");
          return <path key={s.catId} d={path} fill="none" stroke={s.color} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" opacity=".85" />;
        })}
        {/* hover vertical line */}
        {hover !== null && (
          <>
            <line x1={xOf(hover)} x2={xOf(hover)} y1={P} y2={H - PB} stroke="var(--ink-faint)" strokeWidth=".8" strokeDasharray="3,3" />
            {series.map(s => (
              <circle key={s.catId} cx={xOf(hover)} cy={yOf(s.pts[hover].val)} r="4" fill={s.color} stroke="var(--paper-solid)" strokeWidth="2" />
            ))}
          </>
        )}
      </svg>
      {/* hover tooltip */}
      {hover !== null && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: "4px 12px", fontSize: 12, color: "var(--ink-soft)", marginTop: 4 }}>
          <span style={{ fontWeight: 600, color: "var(--ink)" }}>{allDays[hover]}</span>
          {series.map(s => (
            <span key={s.catId} style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <span style={{ width: 8, height: 8, borderRadius: "50%", background: s.color, display: "inline-block" }} />
              {s.emoji} {fmtBare(s.pts[hover].val)}
            </span>
          ))}
        </div>
      )}
      {/* legend */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: "6px 14px", marginTop: 10 }}>
        {series.map(s => (
          <div key={s.catId} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: "var(--ink-soft)" }}>
            <span style={{ width: 10, height: 3, borderRadius: 2, background: s.color, display: "inline-block" }} />
            {s.emoji} {s.name}
          </div>
        ))}
      </div>
    </div>
  );
}

function LineChart({ points, area: showArea = true, color: forcedColor }) {
  const [hover, setHover] = useState(null); // index hovered
  if (!points || points.length < 2) {
    return <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "20px 0" }}>Datos insuficientes.</div>;
  }
  const W = 600, H = 200, P = 10, PB = 22; // PB: padding inferior para etiquetas
  const vals = points.map((p) => p.val);
  const min = Math.min(...vals);
  const max = Math.max(...vals);
  const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
  const range = max - min || 1;
  const xOf = (i) => P + (i / (points.length - 1)) * (W - P * 2);
  const yOf = (v) => H - PB - ((v - min) / range) * (H - P - PB);
  const path = points.map((p, i) => `${i === 0 ? "M" : "L"}${xOf(i).toFixed(1)},${yOf(p.val).toFixed(1)}`).join(" ");
  const areaPath = `${path} L${xOf(points.length - 1).toFixed(1)},${H - PB} L${P},${H - PB} Z`;
  const last = points[points.length - 1].val;
  const first = points[0].val;
  const up = last >= first;
  const stroke = forcedColor || (up ? "var(--green)" : "var(--coral)");
  const avgY = yOf(avg);

  const maxIdx = vals.indexOf(max);
  const minIdx = vals.indexOf(min);

  const handleMove = (e) => {
    const svg = e.currentTarget;
    const rect = svg.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const x = ((clientX - rect.left) / rect.width) * W;
    const i = Math.round(((x - P) / (W - P * 2)) * (points.length - 1));
    setHover(Math.max(0, Math.min(points.length - 1, i)));
  };

  const hp = hover != null ? points[hover] : null;

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, fontSize: 12, color: "var(--ink-soft)" }}>
        <span>{hp ? hp.date : points[0].date}</span>
        <span className="cc-num" style={{ fontWeight: 700, color: stroke, fontSize: 14 }}>
          {hp ? fmt(hp.val) : fmt(last)}
        </span>
        <span>{hp ? "" : "hoy"}</span>
      </div>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block", touchAction: "none" }}
        onMouseMove={handleMove} onMouseLeave={() => setHover(null)}
        onTouchStart={handleMove} onTouchMove={handleMove} onTouchEnd={() => setHover(null)}>
        <defs>
          <linearGradient id="ccLine" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor={stroke} stopOpacity="0.22" />
            <stop offset="100%" stopColor={stroke} stopOpacity="0" />
          </linearGradient>
        </defs>
        {/* promedio */}
        <line x1={P} y1={avgY} x2={W - P} y2={avgY} stroke="var(--ink-faint)" strokeWidth="1" strokeDasharray="4 4" opacity="0.6" />
        <text x={W - P} y={avgY - 4} textAnchor="end" fontSize="10" fill="var(--ink-faint)">prom {fmtBare(avg)}</text>
        {showArea && <path d={areaPath} fill="url(#ccLine)" />}
        <path d={path} fill="none" stroke={stroke} strokeWidth="2.2" strokeLinejoin="round" strokeLinecap="round" />
        {/* máximo y mínimo */}
        <circle cx={xOf(maxIdx)} cy={yOf(max)} r="3.5" fill="var(--green)" />
        <text x={xOf(maxIdx)} y={yOf(max) - 7} textAnchor="middle" fontSize="10" fontWeight="600" fill="var(--green)">{fmtBare(max)}</text>
        <circle cx={xOf(minIdx)} cy={yOf(min)} r="3.5" fill="var(--coral)" />
        <text x={xOf(minIdx)} y={yOf(min) + 14} textAnchor="middle" fontSize="10" fontWeight="600" fill="var(--coral)">{fmtBare(min)}</text>
        {/* cursor hover */}
        {hp && (
          <>
            <line x1={xOf(hover)} y1={P} x2={xOf(hover)} y2={H - PB} stroke="var(--ink-soft)" strokeWidth="1" opacity="0.4" />
            <circle cx={xOf(hover)} cy={yOf(hp.val)} r="5" fill="#fff" stroke={stroke} strokeWidth="2.5" />
          </>
        )}
      </svg>
    </div>
  );
}

/* --------------------- gráfica de pastel (SVG donut) --------------------- */
const PIE_COLORS = ["#B8482A", "#B0863A", "#2E6F4E", "#7B5E2E", "#9E3F26", "#5C7A4C", "#A07344", "#7A4630"];

function PieChart({ data, total }) {
  const size = 160, cx = size / 2, cy = size / 2, r = 70, ir = 42;
  let acc = 0;
  const slices = data.map((d, i) => {
    const frac = total ? d.amt / total : 0;
    const start = acc; acc += frac;
    return { ...d, start, end: acc, color: PIE_COLORS[i % PIE_COLORS.length] };
  });

  function arc(s, e) {
    const a0 = s * Math.PI * 2 - Math.PI / 2;
    const a1 = e * Math.PI * 2 - Math.PI / 2;
    const large = e - s > 0.5 ? 1 : 0;
    const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
    const xi0 = cx + ir * Math.cos(a0), yi0 = cy + ir * Math.sin(a0);
    const xi1 = cx + ir * Math.cos(a1), yi1 = cy + ir * Math.sin(a1);
    return `M${x0},${y0} A${r},${r} 0 ${large} 1 ${x1},${y1} L${xi1},${yi1} A${ir},${ir} 0 ${large} 0 ${xi0},${yi0} Z`;
  }

  return (
    <svg viewBox={`0 0 ${size} ${size}`} style={{ width: 160, height: 160, flexShrink: 0 }}>
      {slices.map((s, i) =>
        s.end - s.start < 0.001 ? null :
        <path key={i} d={arc(s.start, s.end)} fill={s.color} />
      )}
      <text x={cx} y={cy - 4} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="13" fill="var(--ink-soft)">Total</text>
      <text x={cx} y={cy + 14} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="15" fontWeight="600" fill="var(--ink)">{fmt(total)}</text>
    </svg>
  );
}

/* ---------------------- gráfica de barras (SVG) -------------------------- */
function BarsChart({ bars }) {
  const W = 600, H = 180, P = 18;
  const max = Math.max(1, ...bars.map((b) => Math.max(b.expense, b.income)));
  const groupW = (W - P * 2) / bars.length;
  const barW = groupW * 0.38;
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }}>
      {bars.map((b, i) => {
        const cx = P + groupW * i + groupW / 2;
        const hi = (b.income / max) * (H - P * 2);
        const he = (b.expense / max) * (H - P * 2);
        return (
          <g key={i}>
            <rect x={cx - barW - 2} y={H - P - hi} width={barW} height={hi} fill="var(--green)" rx="3" />
            <rect x={cx + 2} y={H - P - he} width={barW} height={he} fill="var(--coral)" rx="3" />
            <text x={cx} y={H - 4} textAnchor="middle" fontSize="10" fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">
              {monthLabel(b.key).split(" ")[0]}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

/* ============= MiniChart para el chat asistente ========================== */
/* Recibe un payload tipo:
   { kind: "line", title, xLabels:["sem 1",...], series:[{name,color,values:[]}], yLabel? }
   { kind: "bars", title, xLabels, series:[{name,color,values}] }
   { kind: "pie",  title, segments:[{label, value, color?}] }
*/
function MiniChart({ data }) {
  if (!data || !data.kind) return null;
  const title = data.title || "";
  return (
    <div style={{ background: "var(--surface)", border: "1px solid var(--line)", borderRadius: 12, padding: 12, marginTop: 6 }}>
      {title && <div style={{ fontSize: 12, fontWeight: 700, color: "var(--ink-soft)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 8 }}>{title}</div>}
      {data.kind === "line" && <MiniLine series={data.series || []} xLabels={data.xLabels || []} />}
      {data.kind === "bars" && <MiniBars series={data.series || []} xLabels={data.xLabels || []} />}
      {data.kind === "pie" && <MiniPie segments={data.segments || []} />}
      {/* leyenda */}
      {(data.kind === "line" || data.kind === "bars") && data.series && data.series.length > 0 && (
        <div style={{ display: "flex", gap: 14, fontSize: 11.5, color: "var(--ink-soft)", marginTop: 8, flexWrap: "wrap" }}>
          {data.series.map((s, i) => (
            <span key={i}>
              <span style={{ display: "inline-block", width: 9, height: 9, background: s.color || PIE_COLORS[i % PIE_COLORS.length], borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />
              {s.name}
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

/* helper para abreviar montos en ejes ($1.2k, $34k) */
function shortMoney(v) {
  const abs = Math.abs(v);
  const sign = v < 0 ? "-" : "";
  if (abs >= 1_000_000) return `${sign}$${(abs / 1_000_000).toFixed(abs >= 10_000_000 ? 0 : 1)}M`;
  if (abs >= 1_000) return `${sign}$${(abs / 1_000).toFixed(abs >= 10_000 ? 0 : 1)}k`;
  return `${sign}$${Math.round(abs)}`;
}

/* genera ticks "redondos" para el eje Y */
function niceTicks(min, max, count = 4) {
  if (min === max) { min -= 1; max += 1; }
  const range = max - min;
  const rough = range / count;
  const pow = Math.pow(10, Math.floor(Math.log10(rough)));
  const norm = rough / pow;
  let step;
  if (norm < 1.5) step = 1;
  else if (norm < 3) step = 2;
  else if (norm < 7) step = 5;
  else step = 10;
  step *= pow;
  const niceMin = Math.floor(min / step) * step;
  const niceMax = Math.ceil(max / step) * step;
  const ticks = [];
  for (let v = niceMin; v <= niceMax + 0.0001; v += step) ticks.push(v);
  return { ticks, min: niceMin, max: niceMax };
}

function MiniLine({ series, xLabels }) {
  const [hover, setHover] = useState(null); // índice de bucket bajo el cursor
  if (!series.length) return null;
  const W = 600, H = 220, PL = 50, PR = 14, PT = 14, PB = 38;
  const all = series.flatMap((s) => s.values || []);
  if (!all.length) return null;
  const rawMin = Math.min(0, ...all);
  const rawMax = Math.max(...all, 1);
  const { ticks, min, max } = niceTicks(rawMin, rawMax, 4);
  const range = max - min || 1;
  const n = xLabels.length;
  const innerW = W - PL - PR;
  const innerH = H - PT - PB;
  const xOf = (i) => PL + (n > 1 ? (i / (n - 1)) * innerW : innerW / 2);
  const yOf = (v) => PT + innerH - ((v - min) / range) * innerH;
  const colWidth = innerW / Math.max(n, 1);

  return (
    <div style={{ position: "relative", width: "100%" }}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }}
        onMouseLeave={() => setHover(null)}>
        {/* grid horizontal */}
        {ticks.map((tv, i) => (
          <g key={i}>
            <line x1={PL} x2={W - PR} y1={yOf(tv)} y2={yOf(tv)}
              stroke={tv === 0 ? "var(--ink-soft)" : "var(--line)"}
              strokeOpacity={tv === 0 ? 0.5 : 1}
              strokeDasharray={tv === 0 ? "none" : "3 4"} />
            <text x={PL - 6} y={yOf(tv) + 3} textAnchor="end" fontSize="10"
              fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">{shortMoney(tv)}</text>
          </g>
        ))}
        {/* etiquetas eje X */}
        {xLabels.map((lbl, i) => (
          <text key={i} x={xOf(i)} y={H - 16} textAnchor="middle" fontSize="11"
            fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">{lbl}</text>
        ))}
        {/* línea guía vertical */}
        {hover != null && (
          <line x1={xOf(hover)} x2={xOf(hover)} y1={PT} y2={PT + innerH}
            stroke="var(--ink-soft)" strokeDasharray="3 3" strokeOpacity="0.5" />
        )}
        {/* series */}
        {series.map((s, si) => {
          const color = s.color || PIE_COLORS[si % PIE_COLORS.length];
          const vals = s.values || [];
          const pts = vals.map((v, i) => `${i === 0 ? "M" : "L"}${xOf(i).toFixed(1)},${yOf(v).toFixed(1)}`).join(" ");
          return (
            <g key={si}>
              <path d={pts} fill="none" stroke={color} strokeWidth="2.4" strokeLinejoin="round" strokeLinecap="round" />
              {vals.map((v, i) => (
                <g key={i}>
                  <circle cx={xOf(i)} cy={yOf(v)} r={hover === i ? 5 : 3.5} fill={color}
                    stroke="var(--surface)" strokeWidth="1.5" />
                </g>
              ))}
            </g>
          );
        })}
        {/* columnas invisibles para captar hover/touch */}
        {xLabels.map((_, i) => {
          const cx = xOf(i);
          return (
            <rect key={i} x={cx - colWidth / 2} y={PT} width={colWidth} height={innerH}
              fill="transparent" style={{ cursor: "pointer" }}
              onMouseEnter={() => setHover(i)}
              onTouchStart={() => setHover(i)} />
          );
        })}
      </svg>
      {/* tooltip */}
      {hover != null && (
        <ChartTooltip
          label={xLabels[hover]}
          rows={series.map((s, si) => ({
            name: s.name, color: s.color || PIE_COLORS[si % PIE_COLORS.length],
            value: (s.values || [])[hover] || 0,
          }))}
          x={(xOf(hover) / W) * 100}
        />
      )}
    </div>
  );
}

function MiniBars({ series, xLabels }) {
  const [hover, setHover] = useState(null);
  if (!series.length) return null;
  const W = 600, H = 220, PL = 50, PR = 14, PT = 14, PB = 38;
  const all = series.flatMap((s) => s.values || []);
  if (!all.length) return null;
  const { ticks, min, max } = niceTicks(0, Math.max(...all, 1), 4);
  const range = max - min || 1;
  const n = xLabels.length;
  const innerW = W - PL - PR;
  const innerH = H - PT - PB;
  const groupW = innerW / Math.max(n, 1);
  const sCount = series.length;
  const barW = Math.min(30, (groupW * 0.7) / sCount);
  const cxOf = (i) => PL + groupW * i + groupW / 2;
  const yOf = (v) => PT + innerH - ((v - min) / range) * innerH;

  return (
    <div style={{ position: "relative", width: "100%" }}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }}
        onMouseLeave={() => setHover(null)}>
        {/* grid horizontal */}
        {ticks.map((tv, i) => (
          <g key={i}>
            <line x1={PL} x2={W - PR} y1={yOf(tv)} y2={yOf(tv)}
              stroke="var(--line)" strokeDasharray={tv === 0 ? "none" : "3 4"} />
            <text x={PL - 6} y={yOf(tv) + 3} textAnchor="end" fontSize="10"
              fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">{shortMoney(tv)}</text>
          </g>
        ))}
        {/* highlight columna activa */}
        {hover != null && (
          <rect x={cxOf(hover) - groupW / 2} y={PT} width={groupW} height={innerH}
            fill="var(--ink)" fillOpacity="0.04" />
        )}
        {/* barras */}
        {xLabels.map((lbl, i) => {
          const cx = cxOf(i);
          return (
            <g key={i}>
              {series.map((s, si) => {
                const color = s.color || PIE_COLORS[si % PIE_COLORS.length];
                const v = (s.values && s.values[i]) || 0;
                const h = (v / max) * innerH;
                const x = cx + (si - (sCount - 1) / 2) * barW - barW / 2;
                return <rect key={si} x={x} y={H - PB - h} width={barW - 2} height={Math.max(h, 0)}
                  fill={color} rx="3" opacity={hover != null && hover !== i ? 0.55 : 1} />;
              })}
              <text x={cx} y={H - 16} textAnchor="middle" fontSize="11"
                fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">{lbl}</text>
            </g>
          );
        })}
        {/* columnas invisibles para hover */}
        {xLabels.map((_, i) => (
          <rect key={i} x={cxOf(i) - groupW / 2} y={PT} width={groupW} height={innerH}
            fill="transparent" style={{ cursor: "pointer" }}
            onMouseEnter={() => setHover(i)}
            onTouchStart={() => setHover(i)} />
        ))}
      </svg>
      {hover != null && (
        <ChartTooltip
          label={xLabels[hover]}
          rows={series.map((s, si) => ({
            name: s.name, color: s.color || PIE_COLORS[si % PIE_COLORS.length],
            value: (s.values || [])[hover] || 0,
          }))}
          x={(cxOf(hover) / W) * 100}
        />
      )}
    </div>
  );
}

/* tooltip flotante posicionado en porcentaje de ancho del contenedor */
function ChartTooltip({ label, rows, x }) {
  // si está a la derecha, anclar a la derecha; si a la izquierda, anclar a la izquierda
  const anchorRight = x > 60;
  return (
    <div style={{
      position: "absolute",
      top: 6,
      [anchorRight ? "right" : "left"]: anchorRight ? `${100 - x + 2}%` : `${x + 2}%`,
      maxWidth: "44%",
      background: "var(--ink)", color: "var(--surface)",
      padding: "8px 11px", borderRadius: 9,
      fontSize: 12, lineHeight: 1.4,
      boxShadow: "0 6px 20px rgba(0,0,0,.25)",
      pointerEvents: "none",
      zIndex: 5,
      whiteSpace: "nowrap",
    }}>
      <div style={{ fontWeight: 700, marginBottom: 4, fontSize: 11, opacity: 0.75, textTransform: "uppercase", letterSpacing: ".05em" }}>{label}</div>
      {rows.map((r, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 7, marginTop: i === 0 ? 0 : 2 }}>
          <span style={{ width: 8, height: 8, borderRadius: 2, background: r.color, flexShrink: 0 }} />
          <span style={{ flex: 1, opacity: 0.8 }}>{r.name}</span>
          <span style={{ fontWeight: 700, fontVariantNumeric: "tabular-nums" }}>{fmt(r.value)}</span>
        </div>
      ))}
    </div>
  );
}

function MiniPie({ segments }) {
  const total = segments.reduce((s, x) => s + (x.value || 0), 0);
  if (total <= 0) return <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>Sin datos.</div>;
  const size = 160, cx = size / 2, cy = size / 2, r = 70, ir = 42;
  let acc = 0;
  const slices = segments.map((seg, i) => {
    const frac = seg.value / total;
    const start = acc; acc += frac;
    return { ...seg, start, end: acc, color: seg.color || PIE_COLORS[i % PIE_COLORS.length] };
  });
  function arc(s, e) {
    const a0 = s * Math.PI * 2 - Math.PI / 2;
    const a1 = e * Math.PI * 2 - Math.PI / 2;
    const large = e - s > 0.5 ? 1 : 0;
    const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
    const xi0 = cx + ir * Math.cos(a0), yi0 = cy + ir * Math.sin(a0);
    const xi1 = cx + ir * Math.cos(a1), yi1 = cy + ir * Math.sin(a1);
    return `M${x0},${y0} A${r},${r} 0 ${large} 1 ${x1},${y1} L${xi1},${yi1} A${ir},${ir} 0 ${large} 0 ${xi0},${yi0} Z`;
  }
  return (
    <div style={{ display: "flex", gap: 14, alignItems: "center", flexWrap: "wrap" }}>
      <svg viewBox={`0 0 ${size} ${size}`} style={{ width: 140, height: 140, flexShrink: 0 }}>
        {slices.map((s, i) => s.end - s.start < 0.001 ? null : <path key={i} d={arc(s.start, s.end)} fill={s.color} />)}
        <text x={cx} y={cy + 5} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="14" fontWeight="600" fill="var(--ink)">{fmt(total)}</text>
      </svg>
      <div style={{ flex: 1, minWidth: 140, display: "flex", flexDirection: "column", gap: 5 }}>
        {slices.slice(0, 6).map((s, i) => {
          const pct = Math.round((s.value / total) * 100);
          return (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5 }}>
              <span style={{ width: 10, height: 10, borderRadius: 2, background: s.color, flexShrink: 0 }} />
              <span style={{ flex: 1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{s.label}</span>
              <span className="cc-num" style={{ fontWeight: 700 }}>{pct}%</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ===================== MODAL: IMPORTAR DESDE EXCEL ======================= */
function ExcelImportModal({ config, txs, onClose, onSave }) {
  const [file, setFile] = useState(null);
  const [rows, setRows] = useState([]); // todas las filas crudas
  const [defaultAccountId, setDefaultAccountId] = useState(
    config.accounts.length === 1 ? config.accounts[0].id : ""
  );
  const [phase, setPhase] = useState("upload"); // upload | processing | review
  const [drafts, setDrafts] = useState([]);
  const [error, setError] = useState(null);
  const inputRef = useRef(null);

  async function handleFile(f) {
    if (!f) return;
    setError(null);
    setFile(f);
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
      const filtered = json.filter((r) => Array.isArray(r) && r.some((c) => c !== "" && c != null));
      if (filtered.length < 2) {
        setError("El archivo está vacío o no tiene suficientes filas.");
        return;
      }
      setRows(filtered);
    } catch (e) {
      console.error(e);
      setError("No pude leer el archivo. ¿Es un Excel válido (.xlsx, .xls, .csv)?");
    }
  }

  async function process() {
    if (!rows.length || !defaultAccountId) return;
    setPhase("processing");
    setError(null);
    try {
      // mandar sample a Claude para detectar columnas
      const sample = rows.slice(0, Math.min(15, rows.length));
      const cats = config.categories
        .filter((c) => c.accountId === defaultAccountId)
        .map((c) => ({ name: c.name, type: c.type }));

      const sys = `Analiza una hoja de cálculo de movimientos bancarios y dime qué columna es cada cosa.
Te paso las primeras filas como arreglo de arreglos. La fila 0 puede ser encabezado (o no).

Devuelve UN SOLO objeto JSON, sin markdown:
{
  "headerRow": <índice de la fila de encabezado, o null si no hay>,
  "columns": {
    "date": <índice de columna de fecha, o null>,
    "description": <índice de columna de concepto/descripción, o null>,
    "amount": <índice de columna de monto, o null>,
    "type": <índice de columna que indica tipo (gasto/ingreso/cargo/abono), o null>,
    "expenseColumn": <índice si hay UNA columna SOLO para gastos (cargo/débito), o null>,
    "incomeColumn": <índice si hay UNA columna SOLO para ingresos (abono/crédito), o null>
  },
  "amountSign": "negative-is-expense" | "positive-is-expense" | "separate-columns" | "type-column",
  "dateFormat": "<descripción breve del formato observado, ej: DD/MM/YYYY>"
}

REGLAS:
- Si hay UNA sola columna de monto y los gastos son negativos → "negative-is-expense"
- Si hay UNA sola columna y todos son positivos pero una columna aparte dice tipo (gasto/ingreso) → "type-column"
- Si hay DOS columnas separadas (una para cargos y otra para abonos) → "separate-columns" + llena expenseColumn e incomeColumn
- Si no estás seguro, usa null en los campos.`;

      const userMsg = `Filas: ${JSON.stringify(sample)}`;
      const raw = await callClaude(sys, [{ role: "user", content: userMsg }]);
      const mapping = parseJSON(raw);

      // construir drafts
      const headerIdx = mapping.headerRow != null ? mapping.headerRow : -1;
      const dataRows = rows.filter((_, i) => i !== headerIdx);
      const sign = mapping.amountSign;
      const cols = mapping.columns || {};

      const ds = [];
      dataRows.forEach((r, idx) => {
        if (!Array.isArray(r)) return;
        const dateRaw = cols.date != null ? r[cols.date] : "";
        const desc = cols.description != null ? String(r[cols.description] || "") : "";

        let amount = 0, type = "expense";
        if (sign === "separate-columns" && cols.expenseColumn != null && cols.incomeColumn != null) {
          const e = parseAmount(r[cols.expenseColumn]);
          const inc = parseAmount(r[cols.incomeColumn]);
          if (inc > 0) { amount = inc; type = "income"; }
          else if (e > 0) { amount = e; type = "expense"; }
        } else if (sign === "type-column" && cols.amount != null && cols.type != null) {
          amount = Math.abs(parseAmount(r[cols.amount]));
          const tt = String(r[cols.type] || "").toLowerCase();
          type = /ingres|abono|cr[eé]dito|depo/.test(tt) ? "income" : "expense";
        } else if (cols.amount != null) {
          const a = parseAmount(r[cols.amount]);
          if (sign === "positive-is-expense") {
            amount = Math.abs(a); type = a >= 0 ? "expense" : "income";
          } else {
            amount = Math.abs(a); type = a < 0 ? "expense" : "income";
          }
        }
        if (amount <= 0) return; // sin monto, descartar

        const date = parseDate(dateRaw) || today();
        // auto-categorizar con keywords locales
        const nd = norm(desc);
        let catId = "";
        const accCats = config.categories.filter((c) => c.accountId === defaultAccountId && c.type === type);
        for (const c of accCats) {
          if ((c.keywords || []).some((kw) => nd.includes(norm(kw)))) { catId = c.id; break; }
        }
        const candidate = { date, amount, type, description: desc.trim(), accountId: defaultAccountId };
        const dup = findDuplicate(candidate, txs || []);
        ds.push({
          tempId: "x" + idx, ...candidate, categoryId: catId,
          selected: !dup,
          duplicate: dup ? { date: dup.date, amount: dup.amount, description: dup.description } : null,
        });
      });

      if (!ds.length) {
        setError("No detecté movimientos válidos. Revisa que el archivo tenga columnas de fecha y monto.");
        setPhase("upload");
        return;
      }
      setDrafts(ds);
      setPhase("review");
    } catch (e) {
      console.error(e);
      setError("Hubo un problema procesando el archivo. Inténtalo de nuevo.");
      setPhase("upload");
    }
  }

  function updateDraft(tempId, patch) {
    setDrafts((prev) => prev.map((d) => (d.tempId === tempId ? { ...d, ...patch } : d)));
  }

  function saveAll() {
    const valid = drafts.filter((d) => d.selected && d.amount > 0);
    if (!valid.length) return;
    let cats = config.categories;
    valid.forEach((d) => {
      if (!d.categoryId || !d.description) return;
      const target = cats.find((c) => c.id === d.categoryId);
      if (!target) return;
      const kws = extractKW(d.description).slice(0, 3);
      if (!kws.length) return;
      const sameName = norm(target.name).trim();
      cats = cats.map((c) =>
        norm(c.name).trim() === sameName && c.type === target.type
          ? { ...c, keywords: [...new Set([...(c.keywords || []), ...kws])] }
          : c
      );
    });
    const txs = valid.map((d) => ({
      id: uid(), type: d.type, amount: d.amount, description: d.description,
      categoryId: d.categoryId || null, accountId: d.accountId, date: d.date,
    }));
    onSave(txs, cats);
  }

  /* ---------- upload ---------- */
  if (phase === "upload") {
    return (
      <div className="cc-overlay" onClick={onClose}>
        <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
          <div className="cc-grip" />
          <div className="cc-sheet-top">
            <h2>Importar desde Excel</h2>
            <button className="cc-sheet-close" onClick={onClose}>×</button>
          </div>
          <p style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 16 }}>
            Sube un archivo de Excel (.xlsx, .xls) o CSV con tus movimientos. La IA detecta las columnas y los importa.
          </p>

          {config.accounts.length > 1 && (
            <div style={{ marginBottom: 14 }}>
              <label className="cc-label">Cuenta destino</label>
              <select className="cc-select" value={defaultAccountId} onChange={(e) => setDefaultAccountId(e.target.value)}
                style={{ borderColor: !defaultAccountId ? "var(--gold)" : "var(--line)" }}>
                <option value="">Elegir cuenta…</option>
                {config.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
            </div>
          )}

          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={(e) => handleFile(e.target.files?.[0])} />

          <button className="cc-card" onClick={() => inputRef.current?.click()}
            style={{ width: "100%", padding: 26, textAlign: "center", border: "2px dashed var(--line)",
              background: "var(--surface-2)", cursor: "pointer", marginBottom: 14 }}>
            <div style={{ fontSize: 30, marginBottom: 6 }}>📄</div>
            <div style={{ fontWeight: 600, fontSize: 14 }}>
              {file ? file.name : "Toca para elegir tu archivo"}
            </div>
            <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 4 }}>
              {file ? `${rows.length} filas detectadas` : "Acepta .xlsx, .xls y .csv"}
            </div>
          </button>

          {error && (
            <div style={{ padding: 11, borderRadius: 10, background: "var(--coral-soft)", color: "var(--coral)", fontSize: 13, marginBottom: 14, fontWeight: 600 }}>
              {error}
            </div>
          )}

          <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 14, fontSize: 15 }}
            disabled={!file || !rows.length || !defaultAccountId} onClick={process}>
            ✨ Detectar e importar
          </button>
        </div>
      </div>
    );
  }

  /* ---------- processing ---------- */
  if (phase === "processing") {
    return (
      <div className="cc-overlay">
        <div className="cc-sheet" style={{ padding: "40px 18px" }}>
          <div className="cc-grip" />
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <div style={{ fontSize: 36, marginBottom: 14 }}>🔍</div>
            <div className="cc-serif" style={{ fontSize: 20, fontWeight: 600, marginBottom: 6 }}>
              Analizando tu archivo…
            </div>
            <div style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 18 }}>
              La IA está detectando las columnas
            </div>
            <div className="cc-dots"><span /><span /><span /></div>
          </div>
        </div>
      </div>
    );
  }

  /* ---------- review ---------- */
  const accCats = (type) => config.categories.filter((c) => c.accountId === defaultAccountId && c.type === type);

  return (
    <ReviewScreen
      drafts={drafts}
      updateDraft={updateDraft}
      accCats={accCats}
      onBack={() => { setDrafts([]); setPhase("upload"); }}
      onSave={saveAll}
      onClose={onClose}
      sourceLabel="Detecté"
    />
  );
}

/* ----- helpers para Excel ----- */
function parseAmount(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[$\s]/g, "").replace(/,/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  // ISO directo
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  // DD/MM/YYYY o DD-MM-YYYY o DD.MM.YYYY
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    let [, dd, mm, yy] = m;
    if (yy.length === 2) yy = "20" + yy;
    return `${yy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`;
  }
  // intentar Date.parse como último recurso
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return null;
}

/* ============= PANTALLA COMPARTIDA: REVISIÓN DE IMPORTACIÓN ============== */
function ReviewScreen({ drafts, updateDraft, accCats, onBack, onSave, onClose, sourceLabel }) {
  const okCount = drafts.filter((d) => d.selected && d.amount > 0).length;
  const dupCount = drafts.filter((d) => d.duplicate).length;
  const dupSelected = drafts.filter((d) => d.duplicate && d.selected).length;

  const markAllDupsOff = () => {
    drafts.forEach((d) => { if (d.duplicate && d.selected) updateDraft(d.tempId, { selected: false }); });
  };
  const markAllDupsOn = () => {
    drafts.forEach((d) => { if (d.duplicate && !d.selected) updateDraft(d.tempId, { selected: true }); });
  };

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Revisa los movimientos</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 12 }}>
          {sourceLabel} <b>{drafts.length}</b>. Edita lo que esté mal o desmarca los que no quieras importar.
        </p>

        {dupCount > 0 && (
          <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "11px 13px",
            background: "#FFF6E0", border: "1px solid #E8C97A", borderRadius: 12, marginBottom: 14, fontSize: 13 }}>
            <span style={{ fontSize: 17 }}>⚠️</span>
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, marginBottom: 2 }}>
                Detecté {dupCount} posible{dupCount === 1 ? "" : "s"} duplicado{dupCount === 1 ? "" : "s"}
              </div>
              <div style={{ color: "var(--ink-soft)", fontSize: 12 }}>
                Movimientos que ya tienes registrados. Están desmarcados por defecto.
              </div>
            </div>
            <button className="cc-btn" style={{ padding: "5px 10px", fontSize: 11.5, whiteSpace: "nowrap" }}
              onClick={dupSelected === dupCount ? markAllDupsOff : markAllDupsOn}>
              {dupSelected === dupCount ? "Desmarcar todos" : "Marcar todos"}
            </button>
          </div>
        )}

        <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 18 }}>
          {drafts.map((d) => {
            const isDup = !!d.duplicate;
            return (
              <div key={d.tempId} className="cc-card" style={{
                padding: 12, opacity: d.selected ? 1 : 0.5,
                borderColor: isDup ? "#E8C97A" : (!d.categoryId ? "var(--gold)" : "var(--line)"),
                background: isDup ? "#FFFCF3" : "var(--surface)",
              }}>
                {isDup && (
                  <div style={{ fontSize: 11.5, fontWeight: 700, color: "#9A6B16",
                    background: "#FFF0CC", padding: "3px 8px", borderRadius: 6,
                    display: "inline-block", marginBottom: 8, letterSpacing: ".02em" }}>
                    🔁 YA EXISTE · {d.duplicate.date}{d.duplicate.description ? ` · "${d.duplicate.description}"` : ""}
                  </div>
                )}
                <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 9 }}>
                  <input type="checkbox" checked={d.selected}
                    onChange={(e) => updateDraft(d.tempId, { selected: e.target.checked })}
                    style={{ width: 18, height: 18, accentColor: "var(--green)" }} />
                  <input className="cc-input" value={d.description} placeholder="Concepto"
                    onChange={(e) => updateDraft(d.tempId, { description: e.target.value })}
                    style={{ flex: 1, fontSize: 14, padding: "7px 10px", fontWeight: 600 }} />
                  <input className="cc-input cc-num" type="number" value={d.amount}
                    onChange={(e) => updateDraft(d.tempId, { amount: parseFloat(e.target.value) || 0 })}
                    style={{ width: 95, fontSize: 14, padding: "7px 10px", fontWeight: 700,
                      color: d.type === "income" ? "var(--green)" : "var(--coral)" }} />
                </div>
                <div style={{ display: "flex", gap: 7, flexWrap: "wrap" }}>
                  <select className="cc-select" value={d.type}
                    onChange={(e) => updateDraft(d.tempId, { type: e.target.value, categoryId: "" })}
                    style={{ flex: "0 0 100px", fontSize: 13, padding: "6px 9px" }}>
                    <option value="expense">Gasto</option>
                    <option value="income">Ingreso</option>
                  </select>
                  <input className="cc-input" type="date" value={d.date}
                    onChange={(e) => updateDraft(d.tempId, { date: e.target.value })}
                    style={{ flex: "0 0 130px", fontSize: 13, padding: "6px 9px" }} />
                  <select className="cc-select" value={d.categoryId}
                    onChange={(e) => updateDraft(d.tempId, { categoryId: e.target.value })}
                    style={{ flex: 1, fontSize: 13, padding: "6px 9px",
                      borderColor: !d.categoryId ? "var(--gold)" : "var(--line)" }}>
                    <option value="">⚠️ Elige categoría</option>
                    {accCats(d.type).map((c) => (
                      <option key={c.id} value={c.id}>{c.emoji} {c.name}</option>
                    ))}
                  </select>
                </div>
              </div>
            );
          })}
        </div>

        <div style={{ display: "flex", gap: 8 }}>
          <button className="cc-btn" style={{ flex: "0 0 auto", padding: 12 }} onClick={onBack}>← Volver</button>
          <button className="cc-btn cc-btn-green" style={{ flex: 1, padding: 13, fontSize: 14 }}
            disabled={okCount === 0} onClick={onSave}>
            Guardar {okCount > 0 ? `(${okCount})` : ""}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============= MODAL: ELEGIR MOVIMIENTO A VINCULAR ======================= */
function LinkPickerModal({ config, txs, currentType, currentAccountId, excludeIds, initialSelectedIds, onClose, onPick }) {
  const [query, setQuery] = useState("");
  const [onlyOpposite, setOnlyOpposite] = useState(false); // por defecto OFF: permitir ambos tipos en el grupo
  const [onlySameAccount, setOnlySameAccount] = useState(true);
  const [selected, setSelected] = useState(new Set(initialSelectedIds || []));

  const exclude = new Set(excludeIds || []);
  const opposite = currentType === "income" ? "expense" : "income";

  // filtrar
  let list = txs.filter((t) => !exclude.has(t.id));
  if (onlyOpposite) list = list.filter((t) => t.type === opposite);
  if (onlySameAccount && currentAccountId) list = list.filter((t) => t.accountId === currentAccountId);
  if (query.trim()) {
    const q = norm(query.trim());
    list = list.filter((t) => {
      const desc = norm(t.description || "");
      const amt = String(t.amount);
      const cat = config.categories.find((c) => c.id === t.categoryId);
      const catName = cat ? norm(cat.name) : "";
      return desc.includes(q) || amt.includes(q) || catName.includes(q) || t.date.includes(query.trim());
    });
  }
  list = [...list].sort((a, b) => b.date.localeCompare(a.date) || (b.id || "").localeCompare(a.id || ""));

  const accMulti = config.accounts.length > 1;
  const toggle = (id) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  // resumen en vivo del grupo (incluyendo solo el actual)
  const selectedTxs = Array.from(selected).map((id) => txs.find((t) => t.id === id)).filter(Boolean);
  const totalIn = selectedTxs.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const totalOut = selectedTxs.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);

  return (
    <div className="cc-overlay" onClick={onClose} style={{ zIndex: 70 }}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Elegir movimientos</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 12 }}>
          Marca todos los movimientos que forman parte del mismo reembolso o encargo. Puedes elegir varios ingresos y/o varios gastos.
        </p>

        <div style={{ marginBottom: 10 }}>
          <input className="cc-input"
            placeholder="🔍 Buscar por concepto, monto, fecha…"
            value={query} onChange={(e) => setQuery(e.target.value)} autoFocus />
        </div>

        <div style={{ display: "flex", gap: 7, marginBottom: 12, flexWrap: "wrap" }}>
          <button onClick={() => setOnlyOpposite(!onlyOpposite)}
            className="cc-chip"
            style={{
              background: onlyOpposite ? "var(--ink)" : "var(--surface)",
              color: onlyOpposite ? "var(--surface)" : "var(--ink)",
              borderColor: onlyOpposite ? "var(--ink)" : "var(--line)",
            }}>
            {onlyOpposite ? "✓ " : ""}Solo {opposite === "income" ? "ingresos" : "gastos"}
          </button>
          {accMulti && (
            <button onClick={() => setOnlySameAccount(!onlySameAccount)}
              className="cc-chip"
              style={{
                background: onlySameAccount ? "var(--ink)" : "var(--surface)",
                color: onlySameAccount ? "var(--surface)" : "var(--ink)",
                borderColor: onlySameAccount ? "var(--ink)" : "var(--line)",
              }}>
              {onlySameAccount ? "✓ " : ""}Misma cuenta
            </button>
          )}
          <span style={{ marginLeft: "auto", alignSelf: "center", fontSize: 11.5, color: "var(--ink-soft)", whiteSpace: "nowrap" }}>
            {list.length} resultado{list.length === 1 ? "" : "s"}
          </span>
        </div>

        {/* lista con checkboxes */}
        <div style={{ display: "flex", flexDirection: "column", gap: 0,
          maxHeight: selected.size > 0 ? "40vh" : "55vh", overflowY: "auto",
          border: "1px solid var(--line)", borderRadius: 12, background: "var(--surface)", padding: "4px 14px" }}>
          {list.length === 0 ? (
            <div style={{ padding: "26px 4px", textAlign: "center", color: "var(--ink-soft)", fontSize: 14 }}>
              <div style={{ fontSize: 28, marginBottom: 6 }}>🔍</div>
              No hay movimientos que coincidan.
              {query && <div style={{ fontSize: 12, marginTop: 4 }}>Prueba con otros términos o quita los filtros.</div>}
            </div>
          ) : (
            renderGroupedByDay(list).map((entry) =>
              entry.type === "header" ? (
                <div key={`h-${entry.date}`} className="cc-day-sep">
                  {(() => { const p = dayParts(entry.date); return (
                    <>
                      <span className="cc-day-num">{p.num}</span>
                      <span className="cc-day-name">{p.desc}</span>
                    </>
                  ); })()}
                </div>
              ) : (
                <LinkPickRow key={entry.t.id} t={entry.t} config={config}
                  selected={selected.has(entry.t.id)}
                  onPick={() => toggle(entry.t.id)} />
              )
            )
          )}
        </div>

        {/* resumen en vivo + botón confirmar */}
        {selected.size > 0 && (
          <div style={{ marginTop: 12, padding: "10px 12px", background: "#FFF6E0", border: "1px solid #E8C97A",
            borderRadius: 11, fontSize: 12.5 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
              <span>Ingresos seleccionados:</span>
              <span className="cc-num" style={{ fontWeight: 700, color: "var(--green)" }}>{fmt(totalIn)}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span>Gastos seleccionados:</span>
              <span className="cc-num" style={{ fontWeight: 700, color: "var(--coral)" }}>{fmt(totalOut)}</span>
            </div>
          </div>
        )}

        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
          {selected.size > 0 && (
            <button className="cc-btn" style={{ padding: "12px 14px", fontSize: 13 }}
              onClick={() => setSelected(new Set())}>
              Limpiar
            </button>
          )}
          <button className="cc-btn cc-btn-green" style={{ flex: 1, padding: 13, fontSize: 14 }}
            disabled={selected.size === 0}
            onClick={() => onPick(Array.from(selected))}>
            {selected.size === 0 ? "Selecciona al menos uno" : `Vincular ${selected.size} ${selected.size === 1 ? "movimiento" : "movimientos"}`}
          </button>
        </div>
      </div>
    </div>
  );
}

function LinkPickRow({ t, config, selected, onPick }) {
  const c = config.categories.find((x) => x.id === t.categoryId);
  const acc = config.accounts.find((a) => a.id === t.accountId);
  const multi = config.accounts.length > 1;
  return (
    <div
      onClick={onPick}
      style={{ display: "flex", alignItems: "center", gap: 11, padding: "11px 0",
        borderBottom: "1px solid var(--line)", cursor: "pointer",
        background: selected ? "var(--green-soft)" : "transparent",
        margin: "0 -10px", paddingLeft: 10, paddingRight: 10,
        borderRadius: selected ? 8 : 0, transition: ".12s" }}
      onMouseEnter={(e) => { if (!selected) e.currentTarget.style.background = "var(--surface-2)"; }}
      onMouseLeave={(e) => { if (!selected) e.currentTarget.style.background = "transparent"; }}>
      <div style={{ fontSize: 22, width: 30, textAlign: "center" }}>{c ? c.emoji : "❔"}</div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 600, fontSize: 14.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
          {t.description || (c ? c.name : "Movimiento")}
        </div>
        <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>
          {c ? c.name : "Sin categoría"}{multi && acc ? ` · ${acc.name}` : ""}
        </div>
      </div>
      <div className="cc-num" style={{ fontWeight: 700, fontSize: 15, whiteSpace: "nowrap",
        color: t.type === "income" ? "var(--green)" : "var(--coral)" }}>
        {t.type === "income" ? "+" : "−"}{fmt(t.amount).replace("-", "")}
      </div>
      {selected && <span style={{ color: "var(--green)", fontSize: 18, fontWeight: 700 }}>✓</span>}
    </div>
  );
}

/* ============= TARJETA RESUMEN del periodo (pestaña Movimientos) ========= */
function SummaryCard({ filter, totalIn, totalOut, topCatRows, topTotal, config }) {
  const showIn = filter === "all" || filter === "income";
  const showOut = filter === "all" || filter === "expense";
  const net = totalIn - totalOut;
  return (
    <div className="cc-card" style={{ padding: "14px 16px", marginBottom: 12 }}>
      <div style={{ display: "flex", alignItems: "stretch", gap: 12 }}>
        {showIn && (
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 2 }}>
            <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
              textTransform: "uppercase", letterSpacing: ".06em" }}>Ingresos</div>
            <div className="cc-serif cc-num" style={{ fontSize: 20, fontWeight: 500, color: "var(--ink)" }}>{fmt(totalIn)}</div>
          </div>
        )}
        {showOut && (
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 2,
            paddingLeft: showIn ? 12 : 0, borderLeft: showIn ? "1px solid var(--line-soft)" : "none" }}>
            <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
              textTransform: "uppercase", letterSpacing: ".06em" }}>Gastos</div>
            <div className="cc-serif cc-num" style={{ fontSize: 20, fontWeight: 500, color: "var(--coral)" }}>{fmt(totalOut)}</div>
          </div>
        )}
        {filter === "all" && (
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 2,
            paddingLeft: 12, borderLeft: "1px solid var(--line-soft)" }}>
            <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
              textTransform: "uppercase", letterSpacing: ".06em" }}>Flujo neto</div>
            <div className="cc-serif cc-num" style={{ fontSize: 20, fontWeight: 500,
              color: net >= 0 ? "var(--ink)" : "var(--coral)" }}>
              {net >= 0 ? "+" : "−"}{fmt(Math.abs(net))}
            </div>
          </div>
        )}
      </div>

      {topCatRows.length > 0 && (
        <div style={{ marginTop: 10, paddingTop: 10, borderTop: "1px solid var(--line)",
          display: "flex", flexWrap: "wrap", gap: 8, fontSize: 11.5 }}>
          {topCatRows.map(([id, amt]) => {
            const c = config.categories.find((x) => x.id === id);
            const pct = topTotal ? Math.round((amt / topTotal) * 100) : 0;
            return (
              <div key={id} style={{ display: "inline-flex", alignItems: "center", gap: 5,
                background: "var(--surface-2)", padding: "4px 9px", borderRadius: 99 }}>
                <span style={{ fontSize: 13 }}>{c ? c.emoji : "❔"}</span>
                <span style={{ fontWeight: 600 }}>{c ? c.name : "Sin cat"}</span>
                <span style={{ color: "var(--ink-soft)" }}>{pct}%</span>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ============= MODAL DETALLE: drill-down de ingresos/gastos/categoría ==== */
function DetailModal({ config, detail, dateRange, onClose, onEditTx }) {
  // ordenar más recientes primero, agrupar por día
  const list = [...(detail.txs || [])].sort((a, b) => b.date.localeCompare(a.date) || (b.id || "").localeCompare(a.id || ""));
  const count = list.length;

  return (
    <div className="cc-overlay" onClick={onClose} style={{ zIndex: 60 }}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div style={{ display: "flex", alignItems: "flex-start", gap: 12, marginBottom: 14 }}>
          <div className="cc-emoji" style={{ fontSize: 32, lineHeight: 1, marginTop: 2 }}>{detail.emoji}</div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <h2 className="cc-serif" style={{ fontSize: 21, fontWeight: 600, marginBottom: 2 }}>{detail.label}</h2>
            <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>
              {rangeLabel(dateRange)} · {count} movimiento{count === 1 ? "" : "s"}
            </div>
          </div>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>

        {/* total */}
        <div className="cc-card" style={{ padding: "14px 16px", marginBottom: 14,
          background: detail.kind === "income" ? "var(--green-soft)" : detail.kind === "expense" ? "var(--coral-soft)" : "var(--surface)" }}>
          <div className="cc-label" style={{ marginBottom: 3 }}>Total</div>
          <div className="cc-num" style={{ fontSize: 28, fontWeight: 700, color: detail.color, letterSpacing: "-.01em" }}>
            {fmt(detail.total)}
          </div>
        </div>

        {/* lista de movimientos */}
        <div style={{ maxHeight: "55vh", overflowY: "auto", border: "1px solid var(--line)", borderRadius: 12,
          background: "var(--surface)", padding: "4px 14px" }}>
          {list.length === 0 ? (
            <div style={{ padding: "26px 4px", textAlign: "center", color: "var(--ink-soft)", fontSize: 14 }}>
              Sin movimientos en este filtro.
            </div>
          ) : (
            renderGroupedByDay(list).map((entry) =>
              entry.type === "header" ? (
                <div key={`h-${entry.date}`} className="cc-day-sep">
                  {(() => { const p = dayParts(entry.date); return (
                    <>
                      <span className="cc-day-num">{p.num}</span>
                      <span className="cc-day-name">{p.desc}</span>
                    </>
                  ); })()}
                  <div className="cc-day-totals">
                    {entry.income > 0 && <span className="pos">+{fmtBare(entry.income)}</span>}
                    {entry.expense > 0 && <span className="neg">−{fmtBare(entry.expense)}</span>}
                  </div>
                </div>
              ) : (
                <TxRow key={entry.t.id} t={entry.t} config={config}
                  onEdit={entry.t.id && !entry.t.id.startsWith("synth-") ? onEditTx : null} />
              )
            )
          )}
        </div>
      </div>
    </div>
  );
}
