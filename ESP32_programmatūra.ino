/*
Piezoelektriskā lietus sensora programmatūra
ZPD autors: Eduards Dāvis Ziemelis, 2026
Jelgavas Valsts ģimnāzija

Šis kods tika pilnībā ģenerēts ar mākslīgā intelekta rīku Claude Sonnet 4.5 
(Anthropic), pamatojoties uz darba autora detalizētiem tehniskajiem 
aprakstiem par aparatūras konfigurāciju, datu apstrādes prasībām un 
vēlamajām funkcijām.
*/

// ESP32-C3 
// Compatible with ESP32 Arduino Core 2.x and 3.x

#define MIC_PIN 0                // GPIO0 (ADC0)
#define SAMPLE_RATE 10000        // 10kHz sampling rate
#define BATCH_SIZE 50            // Larger batch = more efficient USB transfer
#define BUFFER_SIZE (BATCH_SIZE * 2)  // Double buffer for safety

uint16_t audio_buffer[BUFFER_SIZE];
volatile int write_index = 0;
volatile int samples_ready = 0;
volatile bool buffer_overflow = false;

// Precise timing
hw_timer_t *sampling_timer = NULL;
portMUX_TYPE timerMux = portMUX_INITIALIZER_UNLOCKED;

// Statistics
unsigned long total_samples = 0;
unsigned long dropped_samples = 0;
unsigned long last_stats_time = 0;

// Timer ISR - runs at exactly SAMPLE_RATE Hz
void IRAM_ATTR onTimer() {
    portENTER_CRITICAL_ISR(&timerMux);
    
    // Read ADC as fast as possible
    uint16_t sample = analogRead(MIC_PIN);
    
    // Store in buffer
    if (write_index < BUFFER_SIZE) {
        audio_buffer[write_index] = sample;
        write_index++;
        samples_ready++;
        total_samples++;
    } else {
        buffer_overflow = true;
        dropped_samples++;
    }
    
    portEXIT_CRITICAL_ISR(&timerMux);
}

void setup() {
    Serial.begin(921600);
    delay(2000);
    
    // Configure ADC for maximum performance
    analogReadResolution(12);           // 12-bit resolution (0-4095)
    analogSetAttenuation(ADC_11db);     // 0-3.3V range
    
    // Pre-read ADC to initialize
    for (int i = 0; i < 10; i++) {
        analogRead(MIC_PIN);
        delay(1);
    }
    
    Serial.println("# ESP32-C3 Surface Microphone v2.1");
    Serial.println("# Optimized for accuracy and speed");
    Serial.print("# Sample Rate: ");
    Serial.print(SAMPLE_RATE);
    Serial.println(" Hz");
    Serial.print("# Batch Size: ");
    Serial.println(BATCH_SIZE);
    Serial.print("# ADC Pin: GPIO");
    Serial.println(MIC_PIN);
    Serial.println("# Format: timestamp_us,sample1,sample2,...");
    Serial.println("# Ready");
    
    delay(500);
    
    // Setup hardware timer for precise sampling
    // This works with ESP32 Arduino Core 3.x (new API)
    #if ESP_ARDUINO_VERSION >= ESP_ARDUINO_VERSION_VAL(3, 0, 0)
        sampling_timer = timerBegin(SAMPLE_RATE);  // Direct frequency in Hz
        timerAttachInterrupt(sampling_timer, &onTimer);
        timerAlarm(sampling_timer, 1, true, 0);  // Trigger every 1 count, auto-reload
    #else
        // ESP32 Arduino Core 2.x (old API)
        sampling_timer = timerBegin(0, 80, true);  // Timer 0, prescaler 80 (1MHz), count up
        timerAttachInterrupt(sampling_timer, &onTimer, true);
        timerAlarmWrite(sampling_timer, 1000000 / SAMPLE_RATE, true);  // Alarm every 100us for 10kHz
        timerAlarmEnable(sampling_timer);
    #endif
    
    last_stats_time = millis();
}

void loop() {
    // Check if we have enough samples to send a batch
    if (samples_ready >= BATCH_SIZE) {
        portENTER_CRITICAL(&timerMux);
        
        // Get timestamp
        unsigned long timestamp = micros();
        
        // Copy samples to send (we'll send BATCH_SIZE samples)
        uint16_t send_buffer[BATCH_SIZE];
        for (int i = 0; i < BATCH_SIZE; i++) {
            send_buffer[i] = audio_buffer[i];
        }
        
        // Shift remaining samples down
        int remaining = write_index - BATCH_SIZE;
        if (remaining > 0) {
            memmove(audio_buffer, audio_buffer + BATCH_SIZE, remaining * sizeof(uint16_t));
        }
        write_index = remaining;
        samples_ready = remaining;
        
        bool overflow = buffer_overflow;
        buffer_overflow = false;
        
        portEXIT_CRITICAL(&timerMux);
        
        // Send data
        Serial.print(timestamp);
        for (int i = 0; i < BATCH_SIZE; i++) {
            Serial.print(',');
            Serial.print(send_buffer[i]);
        }
        Serial.println();
        
        // Report overflow if occurred
        if (overflow) {
            Serial.println("# WARNING: Buffer overflow! Data loss occurred.");
        }
    }
    
    // Statistics every 10 seconds
    if (millis() - last_stats_time > 10000) {
        portENTER_CRITICAL(&timerMux);
        unsigned long total = total_samples;
        unsigned long dropped = dropped_samples;
        portEXIT_CRITICAL(&timerMux);
        
        Serial.print("# Stats: ");
        Serial.print(total);
        Serial.print(" samples, ");
        Serial.print(dropped);
        Serial.print(" dropped (");
        Serial.print((float)dropped / total * 100.0, 2);
        Serial.print("%), uptime: ");
        Serial.print(millis() / 1000);
        Serial.println("s");
        
        last_stats_time = millis();
    }
    
    // Small delay to prevent watchdog timeout
    delay(1);
}
