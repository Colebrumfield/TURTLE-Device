#include <SPI.h>
#include <Adafruit_MAX31856.h>

// Define CS pins
#define MAXCS1 10
#define MAXCS2 9
#define MAXCS3 8
#define MAXCS4 7

// Create MAX31856 instances
Adafruit_MAX31856 max1 = Adafruit_MAX31856(MAXCS1);
Adafruit_MAX31856 max2 = Adafruit_MAX31856(MAXCS2);
Adafruit_MAX31856 max3 = Adafruit_MAX31856(MAXCS3);
Adafruit_MAX31856 max4 = Adafruit_MAX31856(MAXCS4);

unsigned long previousMillis = 0;
unsigned long interval = 1000; // Default 1 second

void setup() {
  Serial.begin(9600);
  max1.begin();
  max2.begin();
  max3.begin();
  max4.begin();
  setThermocoupleType(MAX31856_TCTYPE_T); // Default to Type T
}

void loop() {
  if (Serial.available() > 0) {
    String command = Serial.readStringUntil(';');
    handleCommand(command);
  }

  unsigned long currentMillis = millis();
  if (currentMillis - previousMillis >= interval) {
    previousMillis = currentMillis;
    readAndSendTemperatures();
  }
}

void setThermocoupleType(uint8_t type) {
  max1.setThermocoupleType(type);
  max2.setThermocoupleType(type);
  max3.setThermocoupleType(type);
  max4.setThermocoupleType(type);
}

void handleCommand(String command) {
  if (command.startsWith("RATE:")) {
    interval = command.substring(5).toInt() * 1000; // Convert seconds to milliseconds
  } else if (command.startsWith("TYPE:")) {
    char type = command.charAt(5);
    switch(type) {
      case 'K': setThermocoupleType(MAX31856_TCTYPE_K); break;
      case 'J': setThermocoupleType(MAX31856_TCTYPE_J); break;
      case 'T': setThermocoupleType(MAX31856_TCTYPE_T); break;
      case 'E': setThermocoupleType(MAX31856_TCTYPE_E); break;
      case 'N': setThermocoupleType(MAX31856_TCTYPE_N); break;
      case 'S': setThermocoupleType(MAX31856_TCTYPE_S); break;
      case 'R': setThermocoupleType(MAX31856_TCTYPE_R); break;
      case 'B': setThermocoupleType(MAX31856_TCTYPE_B); break;
      default: break; // Default case does nothing
    }
  }
}

void readAndSendTemperatures() {
  Serial.print("STATUS:");
  readAndReportStatus(max1, 1);
  readAndReportStatus(max2, 2);
  readAndReportStatus(max3, 3);
  readAndReportStatus(max4, 4);
  Serial.println(";");
}

void readAndReportStatus(Adafruit_MAX31856 &max, int tcNumber) {
  uint8_t fault = max.readFault();
  if (fault & MAX31856_FAULT_OPEN) {
    Serial.print("T"); Serial.print(tcNumber); Serial.print(":Not Connected,");
  } else {
    float temperature = max.readThermocoupleTemperature();
    Serial.print("T"); Serial.print(tcNumber); Serial.print(":"); Serial.print(temperature); Serial.print(",");
  }
}
