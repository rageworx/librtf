GXX = g++
SRC = rtftest.cpp
OUT = test

CFLAGS += -I../inc
CFLAGS += -mms-bitfields
CFLAGS += -mconsole
LFLAGS += -L../lib
LFLAGS += -lrtf
LFLAGS += -lole32 -loleaut32 -luuid -lcomctl32 -lwsock32 -lm
LFLAGS += -lgdi32 -luser32 -lkernel32
LFLAGS += -lShlwapi -lcomdlg32 -lIPHLPAPI
LFLAGS += -g

all : $(OUT)

clean:
	@rm -rf $(OUT)

$(OUT):
	@$(GXX) $(CFLAGS) $(SRC) $(LFLAGS) -o $@
