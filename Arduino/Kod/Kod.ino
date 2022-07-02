#include <SPI.h>
#include <RFID.h>
#include <Wire.h>
#define RESET_PIN 9
#define SDA_PIN 10
RFID Rfid(SDA_PIN, RESET_PIN);
String id;
void setup() {
    Serial.begin(9600);
    SPI.begin();
    Rfid.init();
}
void loop() {
    if (Rfid.isCard()) {
        Rfid.readCardSerial();
        id = "_";
        id += String(Rfid.serNum[0]);
        id += String(Rfid.serNum[1]);
        id += String(Rfid.serNum[2]);
        id += String(Rfid.serNum[3]);
        id += "_";
        if (id.length() > 10) {
            Serial.println(id);
            delay(500);
        }
    }
}
