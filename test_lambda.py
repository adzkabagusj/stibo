import sys, pathlib
sys.path.insert(0, 'Lambda/map-stibo-inbound-validate-transform-dev/src')
import reebok.lambda_function as handler
dirs = {'recap': pathlib.Path('S3/maa-web-portal-files/templates/reebok/licensed/')}
print(handler.recap_etl._parse_metadata_from_filename(dirs, 'recap'))
