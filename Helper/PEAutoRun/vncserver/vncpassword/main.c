#include <stdio.h>
#include <string.h>
#include "d3des.c"

static unsigned char d3desObfuscationKey[] = {23, 82, 107, 6, 35, 78, 88, 7};

int main(int argc, char **argv)
{
    if (argc != 2)
    {
        printf("Error: Invalid input, you must specify a password to return an encrypted value.\n");
        return 1;
    }

    char *passwd = argv[1];

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
    deskey(d3desObfuscationKey, EN0);
    des(newPasswd, newPasswd);

    // Write the key to standard output
    printf("0x");
    for (unsigned int i = 0; i < 8; i++)
    {
        fprintf(stdout, "%02x", newPasswd[i]);
    }
    printf("\n");

    return 0;
}
