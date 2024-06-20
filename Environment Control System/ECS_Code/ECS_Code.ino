//Lazuardi Fauzan Primaputra
//Warehouse Zone Environment Control System


// Declare pin
const int pot = A0;
const int TMP36Sensor = A1;
const int redRGB = 13;
const int blueRGB = 12;
const int greenRGB = 11;
const int motor = 9;

// Decalare global variable
int potVal = 0;
int coolingThreshold = 0;
int heatingThreshold = 0;
int userInputCooling = 0;
int userInputHeating = 0;
int TMP36Reading;
float tempVoltage;
float tempReading;
bool coolingValidInput = false;
bool heatingValidInput = false;

void setup(){
  // Setup the pins as input or output
  Serial.begin(9600);
  pinMode(redRGB, OUTPUT);
  pinMode(greenRGB, OUTPUT);
  pinMode(blueRGB, OUTPUT);
  pinMode(motor, OUTPUT);
  pinMode(TMP36Sensor, INPUT);

  // Turn on green LED to denote the system is running
  digitalWrite(redRGB, LOW);
  digitalWrite(greenRGB, HIGH);
  digitalWrite(blueRGB, LOW);
  
  // Prompt the user for Cooling and Heating limit
  getUserInputCooling();
  getUserInputHeating();
  
  // Show the input limit for Cooling and Heating
  Serial.print("Lower Cooling Limit: ");
  Serial.println(coolingThreshold);
  Serial.print("Upper Heating Limit: ");
  Serial.println(heatingThreshold); 
}

void loop(){
  motorControl();
  tempReader();
  
  // Show the blue light if the temperature is above the Cooling limit
  if(tempReading > coolingThreshold){
    digitalWrite(redRGB, LOW);
    digitalWrite(blueRGB, HIGH);
    digitalWrite(greenRGB, LOW);
  }
  
  // Show the red light if the temperature is below the Heating limit
  else if(tempReading < heatingThreshold){
    digitalWrite(redRGB, HIGH);
    digitalWrite(blueRGB, LOW);
    digitalWrite(greenRGB, LOW);
  }
  
  // Turn back to green led if the temperature is between the Cooling and Heating limit
  else{
    digitalWrite(redRGB, LOW);
    digitalWrite(greenRGB, HIGH);
    digitalWrite(blueRGB, LOW);
  }
}

void getUserInputCooling(){
  while(!coolingValidInput){
    Serial.println("Please enter an integer greater than or equal to 0 for the lower limit for Cooling");
    
    while(Serial.available() == 0);
    if(Serial.available()){
      userInputCooling = Serial.parseInt();
      
      if(userInputCooling >= 0){
        coolingValidInput = true;
        coolingThreshold = userInputCooling;
        Serial.println("Thank You");
      } else {
        Serial.println("Error");
        coolingValidInput = false;
      }
    }
  }
}

void getUserInputHeating(){
  while(!heatingValidInput){
    Serial.println("Please enter an integer less than or equal to the lower limit of Cooling as the upper limit for Heating");
    
    while(Serial.available() == 0);
    if(Serial.available()){
      userInputHeating = Serial.parseInt();
      
      if(userInputHeating <= coolingThreshold){
        heatingValidInput = true;
        heatingThreshold = userInputHeating;
        Serial.println("Thank You");
      } else {
        Serial.println("Error");
        heatingValidInput = false;
      }
    }
  }
}

void motorControl(){
  potVal = analogRead(pot);
  potVal = map(potVal, 0, 1023, 0, 255);
  analogWrite(motor, potVal);
}

void tempReader(){
  TMP36Reading = analogRead(TMP36Sensor);
  tempVoltage = TMP36Reading * (5000.0 / 1024.0);
  tempReading = (tempVoltage - 500) / 10.0;
}