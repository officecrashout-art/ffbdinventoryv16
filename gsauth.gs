/**
 * FASHION FIZZ BD - SECURE AUTH SYSTEM (HASHED)
 */

// 1. LOGIN CHECK
function checkLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  // Hash the password the user just typed
  const inputHash = _hash(password);

  // Loop through rows (Skip header)
  for (var i = 1; i < data.length; i++) {
    const dbUser = data[i][0];
    const dbPass = data[i][1];

    // Check if Username matches AND Password Hash matches
    // We treat the stored password as a hash now.
    if (dbUser === username && dbPass === inputHash) {
      return { status: 'success', user: username };
    }
  }
  return { status: 'fail', message: 'Invalid Username or Password' };
}

// 2. HELPER: SHA-256 Hashing Function
// Converts "tee404" -> "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8"
function _hash(input) {
  if (!input) return "";
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input.toString());
  
  // Convert byte array to Hex String
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) {
      hashVal += 256;
    }
    if (hashVal.toString(16).length == 1) {
      txtHash += '0';
    }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

/**
 * --- ONE-TIME MIGRATION TOOL ---
 * RUN THIS FUNCTION ONCE INSIDE THE SCRIPT EDITOR.
 * It will convert all your existing plain text passwords in the Sheet to Hashes.
 */
function ADMIN_MIGRATE_PASSWORDS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  // Start from row 2 (Index 1) to skip headers
  for (let i = 1; i < data.length; i++) {
    const currentPass = data[i][1];
    
    // Safety Check: If it looks like a hash (64 chars long), skip it so we don't double-hash
    if (currentPass.length === 64 && /^[0-9a-f]+$/i.test(currentPass)) {
      console.log("Skipping " + data[i][0] + " (Already hashed)");
      continue;
    }

    // Convert to Hash
    const newHash = _hash(currentPass);
    
    // Write back to sheet (Row is i+1 because sheet rows start at 1)
    sheet.getRange(i + 1, 2).setValue(newHash);
    console.log("Secured password for user: " + data[i][0]);
  }
}