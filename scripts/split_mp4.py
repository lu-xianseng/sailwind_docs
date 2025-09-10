import cv2
import os


def extract_frames(video_path, output_folder, frame_rate=30):
    """
    从视频中按指定帧率提取帧

    Args:
        video_path (str): 视频文件路径
        output_folder (str): 输出文件夹路径
        frame_rate (int): 目标帧率，默认30fps
    """
    # 创建输出文件夹
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 打开视频文件
    cap = cv2.VideoCapture(video_path)

    if not cap.isOpened():
        print("无法打开视频文件")
        return

    # 获取视频的基本信息
    fps = cap.get(cv2.CAP_PROP_FPS)  # 视频原始帧率
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    duration = total_frames / fps

    print(f"视频信息: {fps:.2f} FPS, 总帧数: {total_frames}, 时长: {duration:.2f}秒")

    # 计算帧间隔（每隔多少帧取一帧）
    frame_interval = max(1, int(fps / frame_rate))

    frame_count = 0
    saved_count = 0

    while True:
        ret, frame = cap.read()

        if not ret:
            break

        # 按帧间隔保存帧
        if frame_count % frame_interval == 0:
            # 生成输出文件名
            output_path = os.path.join(output_folder, f"frame_{saved_count:06d}.jpg")
            cv2.imwrite(output_path, frame)
            saved_count += 1

        frame_count += 1

    cap.release()
    print(f"分帧完成！共保存 {saved_count} 张图片")
    print(f"输出文件夹: {output_folder}")


# 使用示例
if __name__ == "__main__":
    video_path = "Video_250910152119.mp4"  # 替换为你的视频路径
    output_folder = "frames"  # 输出文件夹名称

    extract_frames(video_path, output_folder, frame_rate=30)