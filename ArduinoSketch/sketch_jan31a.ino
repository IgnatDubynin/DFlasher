#include <Wire.h>
#include <Adafruit_GFX.h>
#include <Adafruit_IS31FL3731.h>

#define SERIAL_BAUD 9600

Adafruit_IS31FL3731 mx;

// Digit patterns 0..9 (8x8)
const uint8_t digitPatterns[10][8] = {
  // 0 
  {0b00111100,
   0b01100110,
   0b01100110,
   0b01100110,
   0b01100110,
   0b01100110,
   0b00111100,
   0b00111100},

  // 1
  {0b00011000,
   0b00111000,
   0b01011000,
   0b00011000,
   0b00011000,
   0b00011000,
   0b00111100,
   0b00111100},

  // 2
  {0b00111100,
   0b01100110,
   0b00000110,
   0b00001100,
   0b00110000,
   0b01100000,
   0b01111110,
   0b01111100},

  // 3
  {0b00111100,
   0b01100110,
   0b00000110,
   0b00011100,
   0b00000110,
   0b01100110,
   0b00111100,
   0b00111100},

  // 4
  {0b00001100,
   0b00011100,
   0b00101100,
   0b01001100,
   0b01001110,
   0b01111110,
   0b00001100,
   0b00001100},

  // 5
  {0b01111110,
   0b01100000,
   0b01100000,
   0b01111100,
   0b00000110,
   0b00000110,
   0b01111110,
   0b01111100},

  // 6 
  {0b01111110,
   0b01100100,
   0b01100000,
   0b01111100,
   0b01100110,
   0b01100110,
   0b00111100,
   0b00111100},

  // 7
  {0b01111110,
   0b01100110,
   0b00000110,
   0b00001100,
   0b00011000,
   0b00011000,
   0b00011000,
   0b00011000},

  // 8
  {0b00111100,
   0b01100110,
   0b01100110,
   0b00111100,
   0b01100110,
   0b01100110,
   0b00111100,
   0b00111100},

  // 9
  {0b00111100,
   0b01100110,
   0b01100110,
   0b00111110,
   0b00000110,
   0b01100110,
   0b00111100,
   0b00011100}
};

// --- состояние ---
uint8_t  currentDigit[2]   = {5, 5};
uint8_t  currentBrightness = 5;     // 0..255
uint16_t currentFreq       = 50;    // Гц (циклы POS<->NEG)
const uint16_t minFreq = 1, maxFreq = 200;

bool          showPositive = true;
unsigned long lastFlipTime = 0;
unsigned long flipInterval = 0;

// пары кадров
const uint8_t FRAME0_POS = 0;
const uint8_t FRAME0_NEG = 1;
const uint8_t FRAME1_POS = 2;
const uint8_t FRAME1_NEG = 3;

uint8_t activePair        = 0;   // 0 или 1
bool    pendingPairSwitch = false;
uint8_t nextPair          = 0;

// parser
char    cmdBuf[8];
uint8_t cmdPos = 0;

// --- helpers frames ---
static inline uint8_t posFrameOf(uint8_t pair) { return pair ? FRAME1_POS : FRAME0_POS; }
static inline uint8_t negFrameOf(uint8_t pair) { return pair ? FRAME1_NEG : FRAME0_NEG; }

static inline void showCurrent() {
  mx.displayFrame(showPositive ? posFrameOf(activePair)
                               : negFrameOf(activePair));
}

// =======================
//  FAST I2C WRITE (BURST)
// =======================
static const uint8_t ISSI_ADDR = 0x74;
static const uint8_t ISSI_CMD  = 0xFD;
static const uint8_t PWM_BASE  = 0x24;

static inline void selectBank(uint8_t bank) {
  Wire.beginTransmission(ISSI_ADDR);
  Wire.write(ISSI_CMD);
  Wire.write(bank);
  Wire.endTransmission();
}

// burst write (bank уже выбран!)
static inline void writeBurstReg(uint8_t reg, const uint8_t *data, uint8_t len) {
  Wire.beginTransmission(ISSI_ADDR);
  Wire.write(reg);
  for (uint8_t i = 0; i < len; i++) Wire.write(data[i]);
  Wire.endTransmission();
}

// clear 144 PWM bytes in 6 bursts (24 bytes each)
static inline void clearFrameFast(uint8_t frame) {
  selectBank(frame);
  for (uint8_t i = 0; i < 6; i++) {
    Wire.beginTransmission(ISSI_ADDR);
    Wire.write((uint8_t)(PWM_BASE + i * 24));
    for (uint8_t j = 0; j < 24; j++) Wire.write((uint8_t)0);
    Wire.endTransmission();
  }
}

// PWM reg for (x,y): 0x24 + (x + y*16)
static inline uint8_t pwmRegXY(uint8_t x, uint8_t y) {
  return (uint8_t)(PWM_BASE + (x + (uint16_t)y * 16u));
}

// ---- FAST buildFramePair ----
static void buildFramePair(uint8_t pair) {
  const uint8_t fp = posFrameOf(pair);
  const uint8_t fn = negFrameOf(pair);

  clearFrameFast(fp);
  clearFrameFast(fn);

  uint8_t row[8];

  // POS
  selectBank(fp);
  for (uint8_t y = 0; y < 8; y++) {
    uint8_t bitsL = digitPatterns[currentDigit[0]][y];
    uint8_t bitsR = digitPatterns[currentDigit[1]][y];

    for (uint8_t x = 0; x < 8; x++) row[x] = (bitsL & (0x80 >> x)) ? currentBrightness : 0;
    writeBurstReg(pwmRegXY(0, y), row, 8);

    for (uint8_t x = 0; x < 8; x++) row[x] = (bitsR & (0x80 >> x)) ? currentBrightness : 0;
    writeBurstReg(pwmRegXY(8, y), row, 8);
  }

  // NEG
  selectBank(fn);
  for (uint8_t y = 0; y < 8; y++) {
    uint8_t bitsL = digitPatterns[currentDigit[0]][y];
    uint8_t bitsR = digitPatterns[currentDigit[1]][y];

    for (uint8_t x = 0; x < 8; x++) row[x] = (bitsL & (0x80 >> x)) ? 0 : currentBrightness;
    writeBurstReg(pwmRegXY(0, y), row, 8);

    for (uint8_t x = 0; x < 8; x++) row[x] = (bitsR & (0x80 >> x)) ? 0 : currentBrightness;
    writeBurstReg(pwmRegXY(8, y), row, 8);
  }
}

// --- frequency ---
static inline void applyFreq(uint16_t f) {
  if (f < minFreq) f = minFreq;
  if (f > maxFreq) f = maxFreq;
  currentFreq  = f;
  flipInterval = 500000UL / currentFreq;
}

static inline void adjustFreq(int d) {
  int f = (int)currentFreq + d;
  applyFreq((uint16_t)f);
}

// --- status output ---
static inline void printStatusLine() {
  Serial.print("D=");
  Serial.print(currentDigit[0]);
  Serial.print(',');
  Serial.print(currentDigit[1]);
  Serial.print(" B=");
  Serial.print(currentBrightness);
  Serial.print(" F=");
  Serial.print(currentFreq);
  Serial.println("Hz");
}

// --- serial processing (1 char per call) ---
static void processSerial() {
  if (!Serial.available()) return;

  char c = (char)Serial.read();

  if (c=='\n' || c=='\r' || c==' ' || c==';') {
    if (cmdPos == 0) return;

    cmdBuf[cmdPos] = 0;

    // Dxy
    if ((cmdBuf[0]=='D' || cmdBuf[0]=='d') && cmdPos==3) {
      uint8_t d0 = (uint8_t)(cmdBuf[1] - '0');
      uint8_t d1 = (uint8_t)(cmdBuf[2] - '0');
      if (d0 < 10 && d1 < 10) {
        currentDigit[0] = d0;
        currentDigit[1] = d1;

        uint8_t newPair = (uint8_t)(activePair ^ 1);
        buildFramePair(newPair);
        nextPair = newPair;
        pendingPairSwitch = true;

        Serial.print("D=");
        Serial.print(d0);
        Serial.print(',');
        Serial.println(d1);
      }
    }
    // Bn
    else if (cmdBuf[0]=='B' || cmdBuf[0]=='b') {
      int b = atoi(cmdBuf + 1);
      if (b < 0) b = 0;
      if (b > 255) b = 255;
      currentBrightness = (uint8_t)b;

      uint8_t newPair = (uint8_t)(activePair ^ 1);
      buildFramePair(newPair);
      nextPair = newPair;
      pendingPairSwitch = true;

      Serial.print("B=");
      Serial.println(currentBrightness);
    }
    // F+ / F- / Fn
    else if ((cmdBuf[0]=='F' || cmdBuf[0]=='f') && cmdPos==2 && cmdBuf[1]=='+') {
      adjustFreq(+1);
      Serial.print("F=");
      Serial.println(currentFreq);
    }
    else if ((cmdBuf[0]=='F' || cmdBuf[0]=='f') && cmdPos==2 && cmdBuf[1]=='-') {
      adjustFreq(-1);
      Serial.print("F=");
      Serial.println(currentFreq);
    }
    else if (cmdBuf[0]=='F' || cmdBuf[0]=='f') {
      int f = atoi(cmdBuf + 1);
      applyFreq((uint16_t)f);
      Serial.print("F=");
      Serial.println(currentFreq);
    }
    // S
    else if (cmdBuf[0]=='S' || cmdBuf[0]=='s') {
      printStatusLine();
    }

    cmdPos = 0;
  } else {
    if (cmdPos < sizeof(cmdBuf) - 1) cmdBuf[cmdPos++] = c;
  }
}

// --- setup ---
void setup() {
  Serial.begin(SERIAL_BAUD);
  delay(50);

  Wire.begin();
  Wire.setClock(400000);

  if (!mx.begin(0x74)) {
    while (1);
  }

  activePair = 0;
  buildFramePair(activePair);

  showPositive = true;
  showCurrent();

  applyFreq(currentFreq);
  lastFlipTime = micros();

  // Стартовый статус (как ты хотел)
  Serial.println("IS31FL3731 panel ready");
  Serial.print("INIT ");
  printStatusLine();
}

// --- loop ---
void loop() {
  unsigned long now = micros();

  if (now - lastFlipTime >= flipInterval) {
    lastFlipTime = now;

    if (pendingPairSwitch) {
      activePair = nextPair;
      pendingPairSwitch = false;
    } else {
      showPositive = !showPositive;
    }
    showCurrent();
  }

  processSerial();
}