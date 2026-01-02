import os
import sys
import json
import pytest
import datetime
import sqlite3
from unittest.mock import MagicMock, patch, mock_open

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from Meteo_Lavoro import (
    SettingsManager, DatabaseManager, MeteoAPI, DataProcessor, 
    PDFReporter, DEFAULT_SETTINGS
)

# --- FIXTURES ---
@pytest.fixture
def mock_settings():
    return DEFAULT_SETTINGS.copy()

@pytest.fixture
def sample_forecast_data():
    items = []
    base_time = datetime.datetime.now().replace(hour=7, minute=0, second=0)
    for i in range(12):
        dt = base_time + datetime.timedelta(hours=3*i)
        items.append({
            'dt': int(dt.timestamp()),
            'main': {'temp': 20.0, 'humidity': 50},
            'wind': {'speed': 5.0},
            'weather': [{'description': 'cielo sereno', 'main': 'Clear'}],
            'rain': {'3h': 0},
            'visibility': 10000
        })
    return {'list': items}

@pytest.fixture
def sample_current_data():
    return {
        'main': {'temp': 18.0, 'humidity': 60},
        'wind': {'speed': 4.0},
        'weather': [{'description': 'poche nuvole', 'main': 'Clouds'}],
        'visibility': 10000,
        'sys': {'sunrise': 1600000000, 'sunset': 1600040000}
    }

# --- TEST SETTINGS ---
def test_settings_load_defaults(tmp_path):
    with patch('Meteo_Lavoro.SETTINGS_FILE', str(tmp_path / "non_existent.json")):
        assert SettingsManager.load() == DEFAULT_SETTINGS

# --- TEST DATA PROCESSOR ---
def test_processor_logic_ok(sample_current_data, sample_forecast_data, mock_settings):
    res = DataProcessor.analyze(sample_current_data, sample_forecast_data, mock_settings['thresholds'])
    assert res['status'] == 'OK'
    assert res['score'] == 100

def test_processor_logic_safety(sample_current_data, sample_forecast_data, mock_settings):
    # Vento Forte
    sample_forecast_data['list'][0]['wind']['speed'] = 25.0
    res = DataProcessor.analyze(sample_current_data, sample_forecast_data, mock_settings['thresholds'])
    assert res['score'] < 100
    assert "Vento" in str(res['alerts'])

# --- TEST API ---
def test_api_success(mock_settings, sample_forecast_data, sample_current_data):
    api = MeteoAPI(mock_settings)
    with patch('requests.get') as mock_get:
        mock_geo = MagicMock(); mock_geo.json.return_value = [{'lat': 10, 'lon': 10, 'name': 'Test', 'country': 'IT'}]
        mock_curr = MagicMock(); mock_curr.json.return_value = sample_current_data
        mock_fore = MagicMock(); mock_fore.json.return_value = sample_forecast_data
        
        mock_get.side_effect = [mock_geo, mock_curr, mock_fore]
        
        curr, fore = api.get_data()
        assert curr == sample_current_data
        assert fore == sample_forecast_data

# --- TEST PDF ---
def test_pdf_generation(tmp_path, sample_current_data, sample_forecast_data, mock_settings):
    summary = DataProcessor.analyze(sample_current_data, sample_forecast_data, mock_settings['thresholds'])
    pdf_path = PDFReporter.generate("TestCity", summary, str(tmp_path))
    assert os.path.exists(pdf_path)
