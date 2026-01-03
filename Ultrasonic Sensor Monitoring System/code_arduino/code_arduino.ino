

#include <Wire.h>
#include <LiquidCrystal_I2C.h>

#define TRIG_PIN 9
#define ECHO_PIN 10
#define RELAIS_CONFORME 7
#define RELAIS_NON_CONFORME 8

LiquidCrystal_I2C lcd(0x27, 16, 2);

float seuil_min = 10.0; // cm
float seuil_max = 30; // cm

int nb_conforme = 0;
int nb_non_conforme = 0;

unsigned long temps_relais = 0;
bool relais_actifs = false;
const unsigned long duree_activation = 3000; // 3 secondes

void setup() {
  Serial.begin(9600);
  pinMode(TRIG_PIN, OUTPUT);
  pinMode(ECHO_PIN, INPUT);
  pinMode(RELAIS_CONFORME, OUTPUT);
  pinMode(RELAIS_NON_CONFORME, OUTPUT);

  digitalWrite(RELAIS_CONFORME, LOW);
  digitalWrite(RELAIS_NON_CONFORME, LOW);

  lcd.init();
  lcd.backlight();
  lcd.setCursor(0, 0);
  lcd.print("ultrasonic_monitoring");
  delay(2000);
  lcd.clear();
}

void loop() {
  float distance = mesurerDistance();
  afficherLCD(distance);
  envoyerSerial(distance);
  verifierConformite(distance);
  gererRelais();
  delay(500);
}

float mesurerDistance() {
  digitalWrite(TRIG_PIN, LOW);
  delayMicroseconds(2);
  digitalWrite(TRIG_PIN, HIGH);
  delayMicroseconds(10);
  digitalWrite(TRIG_PIN, LOW);
  long duree = pulseIn(ECHO_PIN, HIGH);
  float distance = duree * 0.034 / 2;
  return distance;
}

void afficherLCD(float dist) {
  lcd.setCursor(0, 0);
  lcd.print("Dist: ");
  lcd.print(dist);
  lcd.print(" cm   ");
}

void envoyerSerial(float dist) {
  Serial.print("Distance:");
  Serial.print(dist);
  Serial.print("cm, Statut:");
  if (dist >= seuil_min && dist <= seuil_max) {
    Serial.println("Conforme");
  } else {
    Serial.println("Non Conforme");
  }
}

void verifierConformite(float dist) {
  if (!relais_actifs) {
    if (dist >= seuil_min && dist <= seuil_max) {
      lcd.setCursor(0, 1);
      lcd.print("Conforme       ");
      digitalWrite(RELAIS_CONFORME, HIGH);
      digitalWrite(RELAIS_NON_CONFORME, LOW);
      nb_conforme++;
    } else {
      lcd.setCursor(0, 1);
      lcd.print("Non Conforme   ");
      digitalWrite(RELAIS_CONFORME, LOW);
      digitalWrite(RELAIS_NON_CONFORME, HIGH);
      nb_non_conforme++;
    }
    relais_actifs = true;
    temps_relais = millis();
  }
}

void gererRelais() {
  if (relais_actifs && (millis() - temps_relais >= duree_activation)) {
    digitalWrite(RELAIS_CONFORME, LOW);
    digitalWrite(RELAIS_NON_CONFORME, LOW);
    relais_actifs = false;
  }
}

