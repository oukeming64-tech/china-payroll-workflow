"use strict";

const CryptoJS = require("crypto-js");
const { Buffer } = require("buffer");

const HASHERS = {
  md5: CryptoJS.MD5,
  sha1: CryptoJS.SHA1,
  sha256: CryptoJS.SHA256,
  sha384: CryptoJS.SHA384,
  sha512: CryptoJS.SHA512,
};

const HMACS = {
  md5: CryptoJS.HmacMD5,
  sha1: CryptoJS.HmacSHA1,
  sha256: CryptoJS.HmacSHA256,
  sha384: CryptoJS.HmacSHA384,
  sha512: CryptoJS.HmacSHA512,
};

function normalizedAlgorithm(algorithm) {
  return String(algorithm || "")
    .toLowerCase()
    .replaceAll("-", "");
}

function asBuffer(value, encoding) {
  if (Buffer.isBuffer(value)) {
    return value;
  }
  if (typeof value === "string") {
    return Buffer.from(value, encoding);
  }
  if (ArrayBuffer.isView(value)) {
    return Buffer.from(value.buffer, value.byteOffset, value.byteLength);
  }
  if (value instanceof ArrayBuffer) {
    return Buffer.from(value);
  }
  return Buffer.from(value || []);
}

function toWordArray(value, encoding) {
  const bytes = asBuffer(value, encoding);
  return CryptoJS.lib.WordArray.create(
    new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength),
  );
}

function fromWordArray(wordArray) {
  const output = Buffer.alloc(wordArray.sigBytes);
  for (let index = 0; index < wordArray.sigBytes; index += 1) {
    output[index] =
      (wordArray.words[index >>> 2] >>> (24 - (index % 4) * 8)) & 0xff;
  }
  return output;
}

function encoded(buffer, encoding) {
  return encoding ? buffer.toString(encoding) : buffer;
}

function createHash(algorithm) {
  const hasher = HASHERS[normalizedAlgorithm(algorithm)];
  if (!hasher) {
    throw new Error(`Unsupported hash algorithm: ${algorithm}`);
  }
  const chunks = [];
  return {
    update(value, encoding) {
      chunks.push(asBuffer(value, encoding));
      return this;
    },
    digest(encoding) {
      const result = fromWordArray(
        hasher(toWordArray(Buffer.concat(chunks))),
      );
      return encoded(result, encoding);
    },
  };
}

function createHmac(algorithm, key) {
  const hasher = HMACS[normalizedAlgorithm(algorithm)];
  if (!hasher) {
    throw new Error(`Unsupported HMAC algorithm: ${algorithm}`);
  }
  const chunks = [];
  return {
    update(value, encoding) {
      chunks.push(asBuffer(value, encoding));
      return this;
    },
    digest(encoding) {
      const result = fromWordArray(
        hasher(
          toWordArray(Buffer.concat(chunks)),
          toWordArray(key),
        ),
      );
      return encoded(result, encoding);
    },
  };
}

function createAesCipher(encrypt, algorithm, key, iv) {
  const match = /^aes-(128|192|256)-(ecb|cbc)$/i.exec(algorithm);
  if (!match) {
    throw new Error(`Unsupported cipher algorithm: ${algorithm}`);
  }
  const mode = match[2].toLowerCase() === "ecb"
    ? CryptoJS.mode.ECB
    : CryptoJS.mode.CBC;
  let autoPadding = true;
  const chunks = [];
  return {
    setAutoPadding(value) {
      autoPadding = Boolean(value);
      return this;
    },
    update(value, inputEncoding) {
      chunks.push(asBuffer(value, inputEncoding));
      return Buffer.alloc(0);
    },
    final(outputEncoding) {
      const input = toWordArray(Buffer.concat(chunks));
      const options = {
        mode,
        padding: autoPadding
          ? CryptoJS.pad.Pkcs7
          : CryptoJS.pad.NoPadding,
      };
      if (mode === CryptoJS.mode.CBC) {
        options.iv = toWordArray(iv);
      }
      const keyWords = toWordArray(key);
      const result = encrypt
        ? CryptoJS.AES.encrypt(input, keyWords, options).ciphertext
        : CryptoJS.AES.decrypt(
            CryptoJS.lib.CipherParams.create({ ciphertext: input }),
            keyWords,
            options,
          );
      return encoded(fromWordArray(result), outputEncoding);
    },
  };
}

function randomBytes(size) {
  const output = Buffer.alloc(size);
  globalThis.crypto.getRandomValues(output);
  return output;
}

module.exports = {
  createCipheriv(algorithm, key, iv) {
    return createAesCipher(true, algorithm, key, iv);
  },
  createDecipheriv(algorithm, key, iv) {
    return createAesCipher(false, algorithm, key, iv);
  },
  createHash,
  createHmac,
  getHashes() {
    return Object.keys(HASHERS);
  },
  randomBytes,
};
