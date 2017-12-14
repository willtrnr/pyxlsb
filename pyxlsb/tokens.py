# Unary operators
UPLUS =  0x12
UMINUS = 0x13
PERCENT = 0x14

# Binary operators
ADD =    0x03
SUB =    0x04
MUL =    0x05
DIV =    0x06
POWER =  0x07
CONCAT = 0x08
LT =     0x09
LE =     0x0A
EQ =     0x0B
GE =     0x0C
GT =     0x0D
NE =     0x0E
ISECT =  0x0F
LIST =   0x10
RANGE =  0x11

# Function operators (Classified)
FUNC     = 0x21
FUNC_VAR = 0x22
FUNC_CE  = 0x38

# Constant operands
MISS_ARG = 0x16
STR      = 0x17
ERR      = 0x1C
BOOL     = 0x1D
INT      = 0x1E
NUM      = 0x1F

# Constant operands (Classified)
ARRAY    = 0x20

# Operands (Classified)
NAME         = 0x23
REF          = 0x24
AREA         = 0x25
MEM_AREA     = 0x26
MEM_ERR      = 0x27
MEM_NO_MEM   = 0x28
MEM_FUNC     = 0x29
REF_ERR      = 0x2A
AREA_ERR     = 0x2B
REF_N        = 0x2C
AREA_N       = 0x2D
MEM_AREA_N   = 0x2E
MEM_NO_MEM_N = 0x2F
NAME_X       = 0x39
REF_3D       = 0x3A
AREA_3D      = 0x3B
REF_ERR_3D   = 0x3C
AREA_ERR_3D  = 0x3D

# Control
EXP       = 0x01
TBL       = 0x02
PAREN     = 0x15
NLR       = 0x18
ATTR      = 0x19
SHEET     = 0x1A
END_SHEET = 0x1B

# Sizes

sizes = [0] * 0x3F

sizes[STR] = -1 # Variable
sizes[ERR] = 1
sizes[BOOL] = 1
sizes[INT] = 2
sizes[NUM] = 8

sizes[ARRAY] = -1 # Variable

sizes[NAME] = 4
sizes[REF] = 6
sizes[AREA] = 14
sizes[MEM_AREA] = -1 # Variable
sizes[REF_ERR] = sizes[REF]
sizes[AREA_ERR] = sizes[AREA]
sizes[REF_N] = 6
sizes[AREA_N] = 14
sizes[NAME_X] = 6
sizes[REF_3D] = 8
sizes[AREA_3D] = 14
sizes[REF_ERR_3D] = sizes[REF_3D]
sizes[AREA_ERR_3D] = sizes[AREA_ERR_3D]

sizes[EXP] = 6
sizes[TBL] = 6
sizes[ATTR] = 3

sizes[MEM_NO_MEM] = 6
sizes[MEM_FUNC] = 2
sizes[MEM_AREA_N] = 2
sizes[MEM_NO_MEM_N] = 2
sizes[FUNC] = 2
sizes[FUNC_VAR] = 3
sizes[FUNC_CE] = 1
