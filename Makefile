# Makefile for librtf  static build for POSIX
# by Raphael Kim

CPP = gcc
CXX = g++
AR  = ar

INC_PATH = inc
SRC_PATH = src
OBJ_PATH = obj
BIN_PATH = lib
TARGET   = librtf.a

SRCS += $(SRC_PATH)/librtf.cpp
OBJS += $(SRCS:$(SRC_PATH)/%.cpp=$(OBJ_PATH)/%.o)

CFLAGS += -I$(SRC_PATH) -I$(INC_PATH)

all: prepare $(BIN_PATH)/$(TARGET)

prepare:
	@mkdir -p $(OBJ_PATH)
	@mkdir -p $(BIN_PATH)

clean:
	@rm -rf $(OBJ_PATH)/*.o
	@rm -rf $(BIN_PATH)/$(TARGET)

$(OBJS): $(OBJ_PATH)/%.o: $(SRC_PATH)/%.cpp
	@echo "Compiling $< ..."
	@$(CXX) $(CFLAGS) -c $< -o $@

$(BIN_PATH)/$(TARGET): $(OBJS)
	@echo "Linking $@ ..."
	@$(AR) -q $@ $(OBJ_PATH)/*.o
