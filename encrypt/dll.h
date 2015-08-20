#ifndef _DLL_H_
#define _DLL_H_

#if BUILDING_DLL
# define DLLIMPORT __declspec (dllexport)
#else /* Not BUILDING_DLL */
# define DLLIMPORT __declspec (dllimport)
#endif /* Not BUILDING_DLL */

typedef struct SHA1Context
{
    unsigned Message_Digest[5]; /* Message Digest (output)          */

    unsigned Length_Low;        /* Message length in bits           */
    unsigned Length_High;       /* Message length in bits           */

    unsigned char Message_Block[64]; /* 512-bit message blocks      */
    int Message_Block_Index;    /* Index into message block array   */

    int Computed;               /* Is the digest computed?          */
    int Corrupted;              /* Is the message digest corruped?  */
};

#define MAX_PBLOCK_SIZE 18   
#define MAX_SBLOCK_XSIZE 4   
#define MAX_SBLOCK_YSIZE 256
#define MAX_KEY_SIZE 56
typedef struct Blowfish_SBlock
{  
    unsigned int m_uil; /*Hi*/  
    unsigned int m_uir; /*Lo*/  
};  
typedef struct Blowfish
{  
    Blowfish_SBlock m_oChain;  
    unsigned int m_auiP[MAX_PBLOCK_SIZE];  
    unsigned int m_auiS[MAX_SBLOCK_XSIZE][MAX_SBLOCK_YSIZE];  
};  
#endif /* _DLL_H_ */
