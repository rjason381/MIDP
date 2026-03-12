"use strict";

(function attachMidpStorage(globalScope) {
  const DEFAULT_DB_NAME = "midp_web_storage";
  const DEFAULT_STORE_NAME = "kv";

  function canUseLocalStorage() {
    try {
      return typeof globalScope.localStorage !== "undefined" && globalScope.localStorage !== null;
    } catch (error) {
      return false;
    }
  }

  function canUseIndexedDb() {
    try {
      return typeof globalScope.indexedDB !== "undefined" && globalScope.indexedDB !== null;
    } catch (error) {
      return false;
    }
  }

  function createUnavailableAdapter() {
    return Object.freeze({
      kind: "unavailable",
      isAvailable() {
        return false;
      },
      getString(key, fallback = "") {
        void key;
        return fallback;
      },
      setString(key, value) {
        void key;
        void value;
        return false;
      },
      getJson(key, fallback = null) {
        void key;
        return fallback;
      },
      setJson(key, value) {
        void key;
        void value;
        return false;
      },
      remove(key) {
        void key;
        return false;
      },
      async hydrateString(key, fallback = "") {
        void key;
        return fallback;
      },
      async hydrateJson(key, fallback = null) {
        void key;
        return fallback;
      }
    });
  }

  function createLocalStorageAdapter() {
    if (!canUseLocalStorage()) {
      return createUnavailableAdapter();
    }

    const storage = globalScope.localStorage;
    return Object.freeze({
      kind: "localStorage",
      isAvailable() {
        return true;
      },
      getString(key, fallback = "") {
        try {
          const value = storage.getItem(String(key));
          return typeof value === "string" ? value : fallback;
        } catch (error) {
          return fallback;
        }
      },
      setString(key, value) {
        try {
          storage.setItem(String(key), String(value));
          return true;
        } catch (error) {
          return false;
        }
      },
      getJson(key, fallback = null) {
        const raw = this.getString(key, "");
        if (!raw) {
          return fallback;
        }
        try {
          return JSON.parse(raw);
        } catch (error) {
          return fallback;
        }
      },
      setJson(key, value) {
        try {
          storage.setItem(String(key), JSON.stringify(value));
          return true;
        } catch (error) {
          return false;
        }
      },
      remove(key) {
        try {
          storage.removeItem(String(key));
          return true;
        } catch (error) {
          return false;
        }
      },
      async hydrateString(key, fallback = "") {
        return this.getString(key, fallback);
      },
      async hydrateJson(key, fallback = null) {
        return this.getJson(key, fallback);
      }
    });
  }

  function createHybridStorageAdapter(options = {}) {
    const localAdapter = createLocalStorageAdapter();
    if (!canUseIndexedDb()) {
      return localAdapter;
    }

    const dbName = typeof options.dbName === "string" && options.dbName.trim()
      ? options.dbName.trim()
      : DEFAULT_DB_NAME;
    const storeName = typeof options.storeName === "string" && options.storeName.trim()
      ? options.storeName.trim()
      : DEFAULT_STORE_NAME;
    let dbPromise = null;

    function getDb() {
      if (dbPromise) {
        return dbPromise;
      }
      dbPromise = new Promise((resolve, reject) => {
        try {
          const request = globalScope.indexedDB.open(dbName, 1);
          request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(storeName)) {
              db.createObjectStore(storeName);
            }
          };
          request.onsuccess = () => {
            resolve(request.result);
          };
          request.onerror = () => {
            reject(request.error || new Error("indexeddb_open_failed"));
          };
        } catch (error) {
          reject(error);
        }
      }).catch(() => null);
      return dbPromise;
    }

    async function readIndexedDbString(key) {
      const db = await getDb();
      if (!db) {
        return "";
      }
      return new Promise((resolve) => {
        try {
          const transaction = db.transaction(storeName, "readonly");
          const store = transaction.objectStore(storeName);
          const request = store.get(String(key));
          request.onsuccess = () => {
            resolve(typeof request.result === "string" ? request.result : "");
          };
          request.onerror = () => {
            resolve("");
          };
          transaction.onabort = () => {
            resolve("");
          };
        } catch (error) {
          resolve("");
        }
      });
    }

    async function writeIndexedDbString(key, value) {
      const db = await getDb();
      if (!db) {
        return false;
      }
      return new Promise((resolve) => {
        try {
          const transaction = db.transaction(storeName, "readwrite");
          const store = transaction.objectStore(storeName);
          store.put(String(value), String(key));
          transaction.oncomplete = () => {
            resolve(true);
          };
          transaction.onerror = () => {
            resolve(false);
          };
          transaction.onabort = () => {
            resolve(false);
          };
        } catch (error) {
          resolve(false);
        }
      });
    }

    async function removeIndexedDbValue(key) {
      const db = await getDb();
      if (!db) {
        return false;
      }
      return new Promise((resolve) => {
        try {
          const transaction = db.transaction(storeName, "readwrite");
          const store = transaction.objectStore(storeName);
          store.delete(String(key));
          transaction.oncomplete = () => {
            resolve(true);
          };
          transaction.onerror = () => {
            resolve(false);
          };
          transaction.onabort = () => {
            resolve(false);
          };
        } catch (error) {
          resolve(false);
        }
      });
    }

    return Object.freeze({
      kind: "hybrid",
      isAvailable() {
        return localAdapter.isAvailable() || canUseIndexedDb();
      },
      getString(key, fallback = "") {
        return localAdapter.getString(key, fallback);
      },
      setString(key, value) {
        const text = String(value);
        const savedLocal = localAdapter.setString(key, text);
        void writeIndexedDbString(key, text);
        return savedLocal || canUseIndexedDb();
      },
      getJson(key, fallback = null) {
        return localAdapter.getJson(key, fallback);
      },
      setJson(key, value) {
        let raw = "";
        try {
          raw = JSON.stringify(value);
        } catch (error) {
          return false;
        }
        const savedLocal = localAdapter.setString(key, raw);
        void writeIndexedDbString(key, raw);
        return savedLocal || canUseIndexedDb();
      },
      remove(key) {
        const removedLocal = localAdapter.remove(key);
        void removeIndexedDbValue(key);
        return removedLocal || canUseIndexedDb();
      },
      async hydrateString(key, fallback = "") {
        const indexedValue = await readIndexedDbString(key);
        const localValue = localAdapter.getString(key, "");
        if (localValue && indexedValue && localValue !== indexedValue) {
          void writeIndexedDbString(key, localValue);
          return localValue;
        }
        if (indexedValue) {
          if (localValue !== indexedValue) {
            localAdapter.setString(key, indexedValue);
          }
          return indexedValue;
        }
        if (localValue) {
          void writeIndexedDbString(key, localValue);
          return localValue;
        }
        return fallback;
      },
      async hydrateJson(key, fallback = null) {
        const raw = await this.hydrateString(key, "");
        if (!raw) {
          return fallback;
        }
        try {
          return JSON.parse(raw);
        } catch (error) {
          return fallback;
        }
      }
    });
  }

  globalScope.MIDPStorage = Object.freeze({
    createLocalStorageAdapter,
    createHybridStorageAdapter
  });
})(typeof window !== "undefined" ? window : globalThis);
