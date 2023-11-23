# VNC Password Encrypt & Decrypt
This simple tool is to encrypt or decrypt a vnc password. This has been done many times but I didn't find an example that accepts input as a command line argument and outputs the result to the console, this is especially useful for scripting.

This has only been tested on the encryption method used by TightVNC, some VNC software use other variations of the DES algorithm so this tool might not always work.

    Usage: vncpassword [password]

    [password] can a clear text password and will be converted to encrypted hex.
    [password] can be encrypted hex and will be converted to clear text. (18 characters, starting with 0x)
