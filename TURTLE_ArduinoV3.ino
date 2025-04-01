#include <SPI.h>
#include <Adafruit_MAX31856.h>
#include <Wire.h>

// Define CS pins
#define MAXCS1 10
#define MAXCS2 9
#define cryoLift_Address 9  // Replace with actual I2C address

// Create MAX31856 instances
Adafruit_MAX31856 max1 = Adafruit_MAX31856(MAXCS1);
Adafruit_MAX31856 max2 = Adafruit_MAX31856(MAXCS2);

unsigned long previousMillis = 0;
unsigned long interval = 1000; // Default 1 second
bool cryoLiftEnabled = false;  // Track CryoLift state

void setup() {
  // Cryo Lift
  Wire.begin();

  Serial.begin(9600);
  max1.begin();
  max2.begin();

  // Set noise filter to 60 Hz (USA)
  // 50 Hz elsewhere
  max1.setNoiseFilter(MAX31856_NOISE_FILTER_60HZ);
  max2.setNoiseFilter(MAX31856_NOISE_FILTER_60HZ);

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
}

void handleCommand(String command) {
  if (command.startsWith("RATE:")) {
    interval = command.substring(5).toInt() * 1000; // Convert seconds to milliseconds
  } else if (command.startsWith("TYPE:")) {
    char type = command.charAt(5);
    switch (type) {
      case 'K': setThermocoupleType(MAX31856_TCTYPE_K); break;
      case 'J': setThermocoupleType(MAX31856_TCTYPE_J); break;
      case 'T': setThermocoupleType(MAX31856_TCTYPE_T); break;
      case 'E': setThermocoupleType(MAX31856_TCTYPE_E); break;
      case 'N': setThermocoupleType(MAX31856_TCTYPE_N); break;
      case 'S': setThermocoupleType(MAX31856_TCTYPE_S); break;
      case 'R': setThermocoupleType(MAX31856_TCTYPE_R); break;
      case 'B': setThermocoupleType(MAX31856_TCTYPE_B); break;
      default: break;
    }
  } else if (command == "Cryo:ON") {
    cryoLiftEnabled = true;
  } else if (command == "Cryo:OFF") {
    cryoLiftEnabled = false;
  }
}

void readAndSendTemperatures() {
  Serial.print("STATUS:");
  float temp1 = readAndReportStatus(max1, 1);
  float temp2 = readAndReportStatus(max2, 2);
  
  Serial.println(";");

  if (cryoLiftEnabled && temp1 != NAN && temp2 != NAN) { // Send over I2C only if CryoLift is enabled and both temperatures are valid
    float avgTemp = (temp1 + temp2) / 2.0;

    Wire.beginTransmission(cryoLift_Address);
    Wire.print(String(avgTemp));
    Wire.endTransmission();
  }
}

float readAndReportStatus(Adafruit_MAX31856 &max, int tcNumber) {
  uint8_t fault = max.readFault();
  if (fault & MAX31856_FAULT_OPEN) {
    Serial.print("T"); Serial.print(tcNumber); Serial.print(":Not Connected,");
    return NAN; // Return NaN to indicate a failure
  } else {
    float temperature = max.readThermocoupleTemperature();
    Serial.print("T"); Serial.print(tcNumber); Serial.print(":"); Serial.print(temperature); Serial.print(",");
    return temperature; // Return the valid temperature
  }
}

