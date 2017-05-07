#include "Ansiterm.h"
Ansiterm ansi;

int i = 0;


void setup() {
  // put your setup code here, to run once:
  Serial.begin(38400);
  delay(10);
  
  ansi.eraseScreen();
  ansi.setForegroundColor(RED);
  Serial.print("Rood ");
}

void loop() {
  // put your main code here, to run repeatedly:
  ansi.setBackgroundColor(WHITE);
  ansi.setForegroundColor(RED);
  Serial.print("Rood ");
  //delay(500);
  ansi.setBackgroundColor(BLACK);
  ansi.setForegroundColor(GREEN);
  Serial.print("Groe ");
  //delay(500);
  ansi.setBackgroundColor(CYAN);
  ansi.setForegroundColor(BLUE);
  Serial.print("Blau ");
  //delay(500);

  i++;


  if(i > 10){
    i = 0;
    Serial.println();
  }

}




