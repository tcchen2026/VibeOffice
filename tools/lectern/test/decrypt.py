"""Independent check of Lectern's encrypted saves: decrypt an ECMA-376 Agile-encrypted package (the format
PowerPoint 2010+ writes) with Python's `cryptography`, verifying the password verifier and the HMAC.
Usage: python3 -I decrypt.py in.pptx password out.pptx"""
import base64, hashlib, hmac, struct, sys
import xml.etree.ElementTree as ET
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

def cfb_streams(b):
    assert b[:8] == bytes.fromhex('d0cf11e0a1b11ae1'), 'not a compound file'
    ssz = 1 << struct.unpack_from('<H', b, 0x1e)[0]; mssz = 1 << struct.unpack_from('<H', b, 0x20)[0]
    nfat, dirstart = struct.unpack_from('<II', b, 0x2c); cutoff = struct.unpack_from('<I', b, 0x38)[0]
    minifat, nminifat, difat, ndifat = struct.unpack_from('<IIII', b, 0x3c)
    sec = lambda n: b[(n + 1) * ssz:(n + 2) * ssz]
    fatsecs = list(struct.unpack_from('<109I', b, 0x4c))[:nfat]
    d = difat
    while len(fatsecs) < nfat and d < 0xfffffffa:
        s = sec(d); vals = struct.unpack('<%dI' % (ssz // 4), s); fatsecs += list(vals[:-1]); d = vals[-1]
    fat = []
    for f in fatsecs[:nfat]: fat += struct.unpack('<%dI' % (ssz // 4), sec(f))
    def chain(start):
        out = []; n = start
        while n < 0xfffffffa: out.append(n); n = fat[n]
        return out
    data = lambda start: b''.join(sec(n) for n in chain(start))
    dirb = data(dirstart)
    ents = []
    for i in range(len(dirb) // 128):
        e = dirb[i * 128:(i + 1) * 128]
        nl = struct.unpack_from('<H', e, 64)[0]
        name = e[:max(0, nl - 2)].decode('utf-16le')
        ents.append((name, e[66], struct.unpack_from('<I', e, 116)[0], struct.unpack_from('<I', e, 120)[0]))
    root = ents[0]
    ministream = data(root[2]) if root[3] else b''
    mfat = []
    if nminifat:
        mb = data(minifat); mfat = list(struct.unpack('<%dI' % (len(mb) // 4), mb))
    def mini(start, size):
        out = b''; n = start
        while n < 0xfffffffa: out += ministream[n * mssz:(n + 1) * mssz]; n = mfat[n]
        return out[:size]
    res = {}
    for name, t, start, size in ents:
        if t == 2: res[name] = mini(start, size) if size < cutoff else data(start)[:size]
    return res

HASH = {'SHA1': 'sha1', 'SHA256': 'sha256', 'SHA384': 'sha384', 'SHA512': 'sha512'}
def H(alg, *parts):
    h = hashlib.new(alg)
    for p in parts: h.update(p)
    return h.digest()
def fix(x, n, pad): return x[:n] if len(x) >= n else x + bytes([pad]) * (n - len(x))
def aes_dec(key, iv, data): d = Cipher(algorithms.AES(key), modes.CBC(iv)).decryptor(); return d.update(data) + d.finalize()

def main(src, pw, dst):
    st = cfb_streams(open(src, 'rb').read())
    info, pkg = st['EncryptionInfo'], st['EncryptedPackage']
    assert info[:4] == b'\x04\x00\x04\x00', 'not agile encryption'
    x = ET.fromstring(info[8:])
    ns = {'e': 'http://schemas.microsoft.com/office/2006/encryption', 'p': 'http://schemas.microsoft.com/office/2006/keyEncryptor/password'}
    kd = x.find('e:keyData', ns); ek = x.find('.//p:encryptedKey', ns); di = x.find('e:dataIntegrity', ns)
    alg = HASH[ek.get('hashAlgorithm')]; salt = base64.b64decode(ek.get('saltValue')); kb = int(ek.get('keyBits')) // 8
    h = H(alg, salt, pw.encode('utf-16le'))
    for i in range(int(ek.get('spinCount'))): h = H(alg, struct.pack('<I', i), h)
    der = lambda blk: fix(H(alg, h, blk), kb, 0x36)
    vin = aes_dec(der(bytes.fromhex('fea7d2763b4b9e79')), salt, base64.b64decode(ek.get('encryptedVerifierHashInput')))[:int(ek.get('saltSize'))]
    vval = aes_dec(der(bytes.fromhex('d7aa0f6d3061344e')), salt, base64.b64decode(ek.get('encryptedVerifierHashValue')))
    assert H(alg, vin) == vval[:len(H(alg, vin))], 'password verifier mismatch'
    key = aes_dec(der(bytes.fromhex('146e0be7abacd0d6')), salt, base64.b64decode(ek.get('encryptedKeyValue')))[:int(kd.get('keyBits')) // 8]
    kalg = HASH[kd.get('hashAlgorithm')]; ksalt = base64.b64decode(kd.get('saltValue'))
    # data integrity (HMAC over the whole EncryptedPackage stream)
    ivk = fix(H(kalg, ksalt, bytes.fromhex('5fb2ad010cb9e1f6')), 16, 0x36); ivv = fix(H(kalg, ksalt, bytes.fromhex('a0677f02b22c8433')), 16, 0x36)
    hk = aes_dec(key, ivk, base64.b64decode(di.get('encryptedHmacKey')))[:hashlib.new(kalg).digest_size]
    hv = aes_dec(key, ivv, base64.b64decode(di.get('encryptedHmacValue')))[:hashlib.new(kalg).digest_size]
    assert hmac.new(hk, pkg, kalg).digest() == hv, 'HMAC mismatch'
    size = struct.unpack_from('<Q', pkg, 0)[0]; out = b''
    body = pkg[8:]
    for seg in range(0, (len(body) + 4095) // 4096):
        iv = fix(H(kalg, ksalt, struct.pack('<I', seg)), 16, 0x36)
        out += aes_dec(key, iv, body[seg * 4096:(seg + 1) * 4096])
    open(dst, 'wb').write(out[:size])
    print('decrypted', size, 'bytes; verifier and HMAC OK')

if __name__ == '__main__': main(*sys.argv[1:4])
