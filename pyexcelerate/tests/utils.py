import os
import shutil

def get_output_path(fn):
	out_dir = os.path.join(os.path.dirname(__file__), 'output')
	if not os.path.exists(out_dir):
		os.mkdir(out_dir)
	return os.path.join(out_dir, fn)
