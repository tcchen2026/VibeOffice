// Published Apache POI fixture passwords and expected content, pinned at ae62bb51.
// Sources: TestCryptoAPI, TestDocumentEncryption, TestPictures, TestHxxFEncryption,
// TestBugs.bug50833 and test-data/poi-integration-exceptions.csv. No password search.
export const encrypted = [
  ['xls', 'legacy-excel/poi__47d331bcc24c__password.xls', 'password', 'A ZIP bomb is a variant'],
  ['xls', 'legacy-excel/poi__a5e42b36f86b__xor-encryption-abc.xls', 'abc', null],
  ['xls', 'legacy-excel/poi__dae995b60ef4__50833.xls', null, 'test cell value'],
  ['xls', 'legacy-excel/poi__ec8640c9112e__35897-type4.xls', 'freedom', 'hello there!'],
  ['doc', 'legacy-word/poi__7878f791ecc8__password_tika_binaryrc4.doc', 'tika', 'This is an encrypted Word 2007 File.'],
  ['doc', 'legacy-word/poi__926f6261cb01__password_password_cryptoapi.doc', 'password', 'This is a test'],
  ['ppt', 'legacy-powerpoint/poi__1a2980dac63d__ppt_with_png_encrypted.ppt', 'password', null],
  ['ppt', 'legacy-powerpoint/poi__3303dbfa2b1f__Password_Protected-hello.ppt', 'hello', 'This is a protected document'],
  ['ppt', 'legacy-powerpoint/poi__3dd99fd5b0ca__Password_Protected-56-hello.ppt', 'hello', 'This is a protected document'],
  ['ppt', 'legacy-powerpoint/poi__5b09e8b393bf__Password_Protected-np-hello.ppt', 'hello', 'Properties are not protected'],
  ['ppt', 'legacy-powerpoint/poi__67b285afb88f__cryptoapi-proc2356.ppt', 'crypto', 'Dominic Salemno'],
];
