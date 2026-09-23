import torch

# PyTorch版本
print("PyTorch 版本:", torch.__version__)

# CUDA版本
print("CUDA 版本:", torch.version.cuda)

# GPU是否可用
print("CUDA 是否可用:", torch.cuda.is_available())

# 如果显卡可用，打印显卡名字
if torch.cuda.is_available():
    print("GPU显卡名称:", torch.cuda.get_device_name(0))
    print("GPU数量:", torch.cuda.device_count())
else:
    print("当前使用CPU模式")
