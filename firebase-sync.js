import { initializeApp } from "https://www.gstatic.com/firebasejs/12.14.0/firebase-app.js";
import {
  getAuth,
  onAuthStateChanged,
  signInAnonymously
} from "https://www.gstatic.com/firebasejs/12.14.0/firebase-auth.js";
import {
  doc,
  getDoc,
  getFirestore,
  serverTimestamp,
  setDoc
} from "https://www.gstatic.com/firebasejs/12.14.0/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyD58DczJFd3f2LIdgyvKLq9MoiQ7ERtkOw",
  authDomain: "reading-bridge-83d70.firebaseapp.com",
  projectId: "reading-bridge-83d70",
  storageBucket: "reading-bridge-83d70.firebasestorage.app",
  messagingSenderId: "562050185274",
  appId: "1:562050185274:web:ff2a0fd9278afe5e4e1d03"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const FAMILY_DOC_ID = "rb_family_83d70_7f4b9c2e1a6d45a8b0c3f9d2";

let currentUser = null;
let authReady = false;
const authReadyListeners = [];

function familyDoc() {
  if (!currentUser) throw new Error("Cloud sync requires anonymous auth.");
  return doc(db, "families", FAMILY_DOC_ID);
}

function notifyAuthReady() {
  while (authReadyListeners.length) {
    const listener = authReadyListeners.shift();
    listener(currentUser);
  }
}

onAuthStateChanged(auth, (user) => {
  currentUser = user;
  authReady = true;
  notifyAuthReady();
  window.dispatchEvent(new CustomEvent("readingbridge:cloud-auth", {
    detail: { signedIn: Boolean(user), familyDocId: FAMILY_DOC_ID }
  }));
});

signInAnonymously(auth).catch((error) => {
  window.dispatchEvent(new CustomEvent("readingbridge:cloud-error", {
    detail: { message: error?.message || "Anonymous cloud sync failed." }
  }));
});

window.ReadingBridgeCloud = {
  isSignedIn() {
    return Boolean(currentUser);
  },

  familyDocId() {
    return FAMILY_DOC_ID;
  },

  whenAuthReady() {
    if (authReady) return Promise.resolve(currentUser);
    return new Promise((resolve) => authReadyListeners.push(resolve));
  },

  async loadCloudStore() {
    if (!currentUser) return null;
    const snapshot = await getDoc(familyDoc());
    return snapshot.exists() ? snapshot.data().store || null : null;
  },

  async saveCloudStore(store) {
    if (!currentUser) return false;
    await setDoc(familyDoc(), {
      store,
      updatedAt: serverTimestamp()
    }, { merge: true });
    return true;
  }
};
