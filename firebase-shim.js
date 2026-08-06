// Liga a aplicação ao Firebase (Firestore), substituindo o "window.storage" que
// antes era fornecido automaticamente pelo Claude. A aplicação em si (app.js)
// não precisa de saber a diferença — continua a chamar window.storage.get/set
// exatamente como antes.

import { initializeApp, deleteApp } from 'firebase/app';
import { getFirestore, doc, getDoc, setDoc, deleteDoc, collection, getDocs } from 'firebase/firestore';
import { getAuth, createUserWithEmailAndPassword, signOut } from 'firebase/auth';
import { firebaseConfig } from './firebase-config.js';

const fbApp = initializeApp(firebaseConfig);
const db = getFirestore(fbApp);
const COLLECTION = 'gestaoCentrais';

window.storage = {
  get: async (key) => {
    const snap = await getDoc(doc(db, COLLECTION, key));
    if (!snap.exists()) return null;
    return { key, value: snap.data().value, shared: true };
  },
  set: async (key, value) => {
    await setDoc(doc(db, COLLECTION, key), { value, atualizadoEm: new Date().toISOString() });
    return { key, value, shared: true };
  },
  delete: async (key) => {
    await deleteDoc(doc(db, COLLECTION, key));
    return { key, deleted: true, shared: true };
  },
  list: async (prefix) => {
    const snaps = await getDocs(collection(db, COLLECTION));
    const keys = [];
    snaps.forEach((d) => {
      if (!prefix || d.id.startsWith(prefix)) keys.push(d.id);
    });
    return { keys, prefix, shared: true };
  },
};

// Gestão de contas de acesso (Firebase Authentication).
//
// criarUtilizador() cria uma conta nova (email + palavra-passe) SEM encerrar a
// sessão de quem está a usar a aplicação neste momento (normalmente um
// Administrador). Isto é feito através de uma "app" Firebase auxiliar e
// temporária, criada só para este pedido e destruída logo a seguir — a app
// principal (e a sessão de quem está autenticado nela) nunca é tocada.
window.firebaseAuth = {
  criarUtilizador: async (email, palavraPasse) => {
    const nomeAuxiliar = `auxiliar-${Date.now()}`;
    const appAuxiliar = initializeApp(firebaseConfig, nomeAuxiliar);
    try {
      const authAuxiliar = getAuth(appAuxiliar);
      const resultado = await createUserWithEmailAndPassword(authAuxiliar, email, palavraPasse);
      const uid = resultado.user.uid;
      await signOut(authAuxiliar);
      return { uid, email };
    } finally {
      await deleteApp(appAuxiliar);
    }
  },
};
