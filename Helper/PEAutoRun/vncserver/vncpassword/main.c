#include <stdio.h>
#include <string.h>
#include "d3des.c"

static unsigned char vncKey[] = {23, 82, 107, 6, 35, 78, 88, 7};

int main(int argc, char **argv)
{
    // Check for help flag or wrong number of arguments
    if (argc != 2 || strcmp(argv[1], "/?") == 0 || strcmp(argv[1], "--help") == 0 || strcmp(argv[1], "-h") == 0)
    {
        printf("\nUsage: vncpassword [password]\n\n");
        printf("    [password] can a clear text password and will be converted to encrypted hex.\n");
        printf("    [password] can be encrypted hex and will be converted to clear text. (18 characters, starting with 0x)\n");
        return 0;
    }

    char *passwd = argv[1];

    // If passwd is 18 characters long and begins with "0x", then it is already encrypted
    if (strlen(passwd) == 18 && passwd[0] == '0' && passwd[1] == 'x')
    {
        // Decrypt the password
        // Remove the "0x" from the beginning of the string and convert to hex
        passwd += 2;
        unsigned char newPasswd[8 + 1];
        for (unsigned int i = 0; i < 8; i++)
        {
            sscanf(passwd + (i * 2), "%2hhx", &newPasswd[i]);
        }

        // Output the decrypted password
        deskey(vncKey, DE1);
        des(newPasswd, newPasswd);
        printf("%s\n", newPasswd);

        // Zero the memory of newPasswd
        memset(newPasswd, 0, sizeof(newPasswd));
    }
    else
    {
        // Encrypt the password

        // Pad password with nulls
        unsigned char newPasswd[8 + 1];
        for (unsigned int i = 0; i < 8; i++)
        {
            if (i < strlen(passwd))
                newPasswd[i] = passwd[i];
            else
                newPasswd[i] = 0;
        }
        // Zero the memory of passwd
        memset(passwd, 0, sizeof(passwd));

        // Create the obfuscated VNC key
        deskey(vncKey, EN0);
        des(newPasswd, newPasswd);

        // Write the key to standard output
        printf("0x");
        for (unsigned int i = 0; i < 8; i++)
        {
            fprintf(stdout, "%02x", newPasswd[i]);
        }
        printf("\n");
    }

    return 0;
}
