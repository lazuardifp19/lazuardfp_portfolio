/*
Lazuardi Fauzan Primaputra
Motion Detection System
*/

const int zone1PIR = 8;		// declare zone 1 LED, PIR, and Button with each pins
const int zone1LED = 7;
const int zone1Button= 3;

const int zone2PIR = 13;	// declare zone 2 LED, PIR, and Button with each pins
const int zone2LED = 12;
const int zone2Button = 2;

bool zone1Motion = false;	// declare variable for PIR output at corresponding zones
bool zone2Motion = false;	
bool zone1MotionLock = false;	// declare variable to "lock" PIR output at each zones
bool zone2MotionLock = false;
volatile bool buttonPushed1 = false;	// declare interrupt flags for buttons at each zones
volatile bool buttonPushed2 = false;

void setup(){
  Serial.begin(9600);
  // declare each pins for input and output
  pinMode(zone1LED, INPUT);
  pinMode(zone2LED, INPUT);
  pinMode(zone1PIR, INPUT);
  pinMode(zone2PIR, INPUT);
  pinMode(zone1Button, OUTPUT);
  pinMode(zone2Button, OUTPUT);
  
  // make sure that LED at each zones are turned off at the start
  digitalWrite(zone1LED, LOW);
  digitalWrite(zone2LED, LOW);

  // setup interface for user
  Serial.print("Initializing alarm");
  delay(500);
  Serial.print(".");
  delay(500);
  Serial.print(".");
  delay(500);
  Serial.println(".");
  delay(700);
  Serial.println("Alarm initialized");
  
  // declare attachInterrupt() for each zones
  attachInterrupt(digitalPinToInterrupt(zone1Button), buttonISR1, RISING);
  attachInterrupt(digitalPinToInterrupt(zone2Button), buttonISR2, RISING);
}

void loop(){
  zone1Motion = digitalRead(zone1PIR);	// read input from PIR sensor at zone 1
  if(zone1Motion == HIGH){	// "lock" output from PIR sensor so LED will still turns on when there is no motion anymore
    zone1MotionLock = true;
  }
  if(zone1MotionLock == true){
    activateAlarmZone1();
  }
  if(buttonPushed1){
    zone1MotionLock = false;
    deactivateAlarmZone1();
  }

  zone2Motion = digitalRead(zone2PIR);	// read input from PIR sensor at zone 2
  if(zone2Motion == HIGH){	// "lock" output from PIR sensor so LED will still turns on when there is no motion anymore
    zone2MotionLock = true;
  }
  if(zone2MotionLock == true){
    activateAlarmZone2();
  }
  if(buttonPushed2){
    zone2MotionLock = false;
    deactivateAlarmZone2();
  }
}

void activateAlarmZone1(){	// turns on the LED at zone 1
  digitalWrite(zone1LED, HIGH);
  Serial.println("Motion detected at Zone 1");
  delay(2000);
}
void activateAlarmZone2(){	// turns on the LED at zone 2
  digitalWrite(zone2LED, HIGH);
  Serial.println("Motion detected at Zone 2");
  delay(2000);
}
void deactivateAlarmZone1(){	// turns off the LED at zone 1
  digitalWrite(zone1LED, LOW);
  Serial.println("Alarm at zone 1 has been deactivated");
  buttonPushed1 = false;
}
void deactivateAlarmZone2(){	//turns off the LED at zone 2
  digitalWrite(zone2LED, LOW);
  Serial.println("Alarm at zone 2 has been deactivated");
  buttonPushed2 = false;
}
void buttonISR1(){	// interrupt service routine for zone 1 button
  buttonPushed1 = true;
}
void buttonISR2(){	// interrupt service routine for zone 2 button
  buttonPushed2 = true;
}
