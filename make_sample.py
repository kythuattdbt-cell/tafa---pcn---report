"""Create a synthetic demonstration, outside real data/output folders."""
import argparse
import tempfile
from pathlib import Path
from tests.test_report import write_sample

if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory',default=str(Path(tempfile.gettempdir())/'tafa-demo'/'data'))
    args=parser.parse_args()
    target=Path(args.directory)
    if (target/'MAU_GIA_LAP.xlsx').exists():
        parser.error('File mẫu đã tồn tại; chọn thư mục mới.')
    write_sample(target)
    print('Đã tạo MAU_GIA_LAP.xlsx (dữ liệu hoàn toàn giả lập).')
