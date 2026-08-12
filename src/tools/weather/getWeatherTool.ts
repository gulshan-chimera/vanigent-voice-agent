import { ToolDefinition } from "../../core/types";
import { ToolValidationError } from "../../core/errors";

type WeatherInput = {
  city: string;
  units?: "celsius" | "fahrenheit";
};

type WeatherOutput = {
  city: string;
  temperature: number;
  units: "celsius" | "fahrenheit";
  condition: string;
  humidity: number;
  windSpeed: number;
  windUnits: string;
  description: string;
};

// Mock weather data for demonstration purposes
const MOCK_WEATHER: Record<
  string,
  { tempC: number; condition: string; humidity: number; windKmh: number }
> = {
  "new york": { tempC: 22, condition: "Partly Cloudy", humidity: 65, windKmh: 14 },
  london: { tempC: 15, condition: "Overcast", humidity: 80, windKmh: 20 },
  tokyo: { tempC: 28, condition: "Sunny", humidity: 70, windKmh: 8 },
  paris: { tempC: 18, condition: "Light Rain", humidity: 75, windKmh: 12 },
  sydney: { tempC: 20, condition: "Clear", humidity: 55, windKmh: 18 },
  mumbai: { tempC: 32, condition: "Humid", humidity: 85, windKmh: 10 },
  dubai: { tempC: 40, condition: "Sunny", humidity: 30, windKmh: 15 },
  berlin: { tempC: 17, condition: "Cloudy", humidity: 60, windKmh: 22 },
  toronto: { tempC: 19, condition: "Partly Cloudy", humidity: 58, windKmh: 16 },
  singapore: { tempC: 31, condition: "Thunderstorms", humidity: 90, windKmh: 6 },
};

function celsiusToFahrenheit(c: number): number {
  return Math.round((c * 9) / 5 + 32);
}

function kmhToMph(kmh: number): number {
  return Math.round(kmh * 0.621371);
}

export const getWeatherTool: ToolDefinition<WeatherInput, WeatherOutput> = {
  name: "weather.get",
  description:
    "Returns current weather information for a given city including temperature, condition, humidity, and wind speed",

  validate(input: unknown): WeatherInput {
    if (typeof input !== "object" || input === null) {
      throw new ToolValidationError(this.name, "Input must be an object");
    }

    const candidate = input as Partial<WeatherInput>;

    if (typeof candidate.city !== "string" || candidate.city.trim().length === 0) {
      throw new ToolValidationError(this.name, "'city' must be a non-empty string");
    }

    if (
      candidate.units !== undefined &&
      candidate.units !== "celsius" &&
      candidate.units !== "fahrenheit"
    ) {
      throw new ToolValidationError(
        this.name,
        "'units' must be either 'celsius' or 'fahrenheit'"
      );
    }

    return {
      city: candidate.city.trim(),
      units: candidate.units ?? "celsius",
    };
  },

  execute(input: WeatherInput): WeatherOutput {
    const key = input.city.toLowerCase();
    const data = MOCK_WEATHER[key];

    if (!data) {
      // Generate random but deterministic-feeling weather for unknown cities
      const hash = key.split("").reduce((acc, ch) => acc + ch.charCodeAt(0), 0);
      const tempC = 10 + (hash % 30);
      const conditions = ["Sunny", "Cloudy", "Partly Cloudy", "Rainy", "Clear"];
      const condition = conditions[hash % conditions.length];
      const humidity = 30 + (hash % 60);
      const windKmh = 5 + (hash % 25);

      const useFahrenheit = input.units === "fahrenheit";
      return {
        city: input.city,
        temperature: useFahrenheit ? celsiusToFahrenheit(tempC) : tempC,
        units: input.units ?? "celsius",
        condition,
        humidity,
        windSpeed: useFahrenheit ? kmhToMph(windKmh) : windKmh,
        windUnits: useFahrenheit ? "mph" : "km/h",
        description: `${input.city}: ${condition}, ${useFahrenheit ? celsiusToFahrenheit(tempC) : tempC}°${useFahrenheit ? "F" : "C"}`,
      };
    }

    const useFahrenheit = input.units === "fahrenheit";
    return {
      city: input.city,
      temperature: useFahrenheit ? celsiusToFahrenheit(data.tempC) : data.tempC,
      units: input.units ?? "celsius",
      condition: data.condition,
      humidity: data.humidity,
      windSpeed: useFahrenheit ? kmhToMph(data.windKmh) : data.windKmh,
      windUnits: useFahrenheit ? "mph" : "km/h",
      description: `${input.city}: ${data.condition}, ${useFahrenheit ? celsiusToFahrenheit(data.tempC) : data.tempC}°${useFahrenheit ? "F" : "C"}`,
    };
  },
};
