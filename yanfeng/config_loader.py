# -*- coding: utf-8 -*-
import importlib.util
import os

script_dir = os.path.dirname(os.path.abspath(__file__))


def get_config_dir() -> str:
    config_dir_file = os.path.join(script_dir, 'config_DIR.py')
    if not os.path.exists(config_dir_file):
        raise FileNotFoundError(f'未找到配置目录文件: {config_dir_file}')

    spec = importlib.util.spec_from_file_location('config_DIR', config_dir_file)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    if not hasattr(module, 'CONFIG_DIR'):
        raise AttributeError('config_DIR.py 中缺少 CONFIG_DIR')

    config_dir = module.CONFIG_DIR
    if not os.path.isabs(config_dir):
        config_dir = os.path.abspath(os.path.join(script_dir, config_dir))

    return config_dir


def get_config_path() -> str:
    config_dir = get_config_dir()
    config_path = os.path.join(config_dir, 'config.py')
    if not os.path.exists(config_path):
        raise FileNotFoundError(f'未找到 config.py: {config_path}')
    return config_path


def load_config_module():
    config_path = get_config_path()
    spec = importlib.util.spec_from_file_location('config', config_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module
