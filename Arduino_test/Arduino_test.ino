#include "Ansiterm.h"
Ansiterm ansi;

void setup() {
  // put your setup code here, to run once:
  Serial.begin(38400);
}

void loop() {
  // put your main code here, to run repeatedly:
  ansi.setForegroundColor(RED);
  Serial.print("Rood ");
  delay(100);
  ansi.setForegroundColor(GREEN);
  Serial.print("Groe ");
  delay(100);
  ansi.setForegroundColor(BLUE);
  Serial.print("Blau ");
  delay(100);
}
