import { initializeApp } from "firebase/app";
import { indexedDBLocalPersistence, initializeAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyCZTrJTGH8Jh5WBMhMrV39mjKddRj7p78w",
  authDomain: "zafi-524b8.firebaseapp.com",
  projectId: "zafi-524b8",
  storageBucket: "zafi-524b8.firebasestorage.app",
  messagingSenderId: "308516673564",
  appId: "1:308516673564:web:9410954d5fc50fd56667d9"
};

const app = initializeApp(firebaseConfig);

export const auth = initializeAuth(app, {
  persistence: indexedDBLocalPersistence,
});

export const db = getFirestore(app);
