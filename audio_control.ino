const int INPUT_PINS[] = { A0, A1, A2, A3 };
const int NUM_PINS = sizeof(INPUT_PINS) / sizeof(int);
const int MARGIN = 5;

int sliderValues[NUM_PINS];
int oldValues[NUM_PINS];

bool close(int a, int b) {
  return abs(a - b) < MARGIN;
}

bool valuesAreDifferent() {
  for (int i = 0; i < NUM_PINS; i++) {
    if (!close(sliderValues[i], oldValues[i])) {
      return true;
    }
  }
  return false;
}

void collectData() {
  for (int i = 0; i < NUM_PINS; i++) {
    sliderValues[i] = analogRead(INPUT_PINS[i]);
  }
}

String createPinMessage() {
  String output = "";
  for (int i = 0; i < NUM_PINS; i++) {
    output += String(sliderValues[i]);
    if (i < NUM_PINS - 1) {
      output += String("|");
    }
  }
  return output;
}

void clearSerial() {
  while (Serial.read() != -1) {
    Serial.read();
    delay(200);
  }
}


void setup() {
  for (int i = 0; i < NUM_PINS; i++) {
    pinMode(INPUT_PINS[i], INPUT);
  }
  Serial.begin(9600);
  clearSerial();
}

void loop() {
  collectData();
  if (valuesAreDifferent()) {
    Serial.println(createPinMessage());
    for (int i = 0; i < NUM_PINS; i++) {
      oldValues[i] = sliderValues[i];
    }
  }
}
