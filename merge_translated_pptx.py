#!/usr/bin/env python3
"""
merge_translated_pptx.py - 翻译 PPTX 文本与原版高清大文件/视频回填工具

核心设计（Frontend-XML + Backend-Media Layout）:
1. 字符级纯文本精准插槽替换（Byte-Exact Tag Injection）：
   - 绝不通过 DOM 重构全量 XML（规避 ElementTree 自动生成 ns0: 异常前缀导致 Office 假死与修复报错）。
   - 仅对 `<a:t>...</a:t>` 内部文本进行流式序列替换。
2. 物理排布优化（规避 32 位有符号寻址溢出）：
   - 将所有核心 XML 文件（幻灯片、母版、版式、关系、属性等）排在 ZIP 容器最前端（Offset < 10MB）。
   - 将 26 个超大媒体视频（MP4/MOV）排在 ZIP 容器后端，并保持 `compress_type=0` (ZIP_STORED 不压缩)。
   - 如此 PowerPoint 启动时读取 Slide XML 的 Seek 指针严格在 32 位安全区间内，绝不触发“损坏修复”弹窗。
3. 纯标准 ZIP32 格式输出，不注入任何异常 ZIP64 结构。
"""

import argparse
import os
import re
import struct
import sys
import zlib
import zipfile


def _inject_translated_text(orig_xml_bytes: bytes, trans_xml_bytes: bytes) -> bytes:
    """将翻译后的 <a:t> 文本序列注入到原版 XML 中，保留原版全部标签结构。"""
    orig_xml = orig_xml_bytes.decode('utf-8')
    trans_xml = trans_xml_bytes.decode('utf-8')
    trans_texts = re.findall(r'<a:t\b[^>]*>(.*?)</a:t>', trans_xml, flags=re.DOTALL)
    idx = [0]
    def replace_tag(match):
        attr = match.group(1)
        if idx[0] < len(trans_texts):
            new_text = trans_texts[idx[0]]
            idx[0] += 1
            return f'<a:t{attr}>{new_text}</a:t>'
        return match.group(0)
    new_xml = re.sub(r'<a:t(\b[^>]*)>.*?</a:t>', replace_tag, orig_xml, flags=re.DOTALL)
    return new_xml.encode('utf-8')


def _build_local_file_entry(filename: bytes, data: bytes, compress: bool,
                            flags: int = 0, mod_time: int = 0, mod_date: int = 0x0021) -> tuple:
    """
    构建 ZIP Local File Entry（Local File Header + 数据）。
    返回 (entry_bytes, compressed_size, uncompressed_size, crc32, compress_method)
    """
    crc = zlib.crc32(data) & 0xFFFFFFFF
    uncompressed_size = len(data)

    if compress:
        co = zlib.compressobj(6, zlib.DEFLATED, -15)  # raw deflate
        compressed = co.compress(data) + co.flush()
        compress_method = 8  # DEFLATED
        flags = 6 if flags == 0 else flags
    else:
        compressed = data
        compress_method = 0  # STORED
        flags = 0

    compressed_size = len(compressed)

    lfh = struct.pack(
        '<4sHHHHHIIIHH',
        b'\x50\x4b\x03\x04',
        20,                      # Version needed
        flags,
        compress_method,
        mod_time,
        mod_date,
        crc,
        compressed_size,
        uncompressed_size,
        len(filename),
        0,                       # Extra length = 0
    )
    entry = lfh + filename + compressed
    return entry, compressed_size, uncompressed_size, crc, compress_method, flags


def _build_cd_entry(filename: bytes, compress_method: int, flags: int,
                    mod_time: int, mod_date: int, crc: int,
                    compressed_size: int, uncompressed_size: int,
                    local_header_offset: int, external_attr: int = 0) -> bytes:
    """构建 Central Directory File Header。"""
    return struct.pack(
        '<4sHHHHHHIIIHHHHHII',
        b'\x50\x4b\x01\x02',
        20,
        20,
        flags,
        compress_method,
        mod_time,
        mod_date,
        crc,
        compressed_size,
        uncompressed_size,
        len(filename),
        0,
        0,
        0,
        0,
        external_attr,
        local_header_offset,
    ) + filename


def _build_eocd(num_entries: int, cd_size: int, cd_offset: int) -> bytes:
    """构建 End-of-Central-Directory Record。"""
    return struct.pack(
        '<4sHHHHIIH',
        b'\x50\x4b\x05\x06',
        0,
        0,
        num_entries,
        num_entries,
        cd_size,
        cd_offset,
        0,
    )


def transplant_text(orig_pptx: str, trans_pptx: str, out_pptx: str):
    if not os.path.exists(orig_pptx):
        print(f"Error: Original PPTX not found: {orig_pptx}", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(trans_pptx):
        print(f"Error: Translated PPTX not found: {trans_pptx}", file=sys.stderr)
        sys.exit(1)

    print(f"[*] PPTX 回填与结构优化流水线启动 ...")
    print(f"    - 原版底稿 (含视频): {orig_pptx}")
    print(f"    - DeepL 译本 (含文字): {trans_pptx}")

    z_orig = zipfile.ZipFile(orig_pptx, 'r')
    z_trans = zipfile.ZipFile(trans_pptx, 'r')

    trans_slide_names = {n for n in z_trans.namelist()
                         if n.startswith('ppt/slides/slide') and n.endswith('.xml')}
    orig_infolist = z_orig.infolist()

    # 1. 分离条目：XML 结构文件（前置）与 Media 媒体文件（后置）
    xml_entries = []
    media_entries = []

    for info in orig_infolist:
        if info.filename.startswith('ppt/media/'):
            media_entries.append(info)
        else:
            xml_entries.append(info)

    print(f"    - 核心结构/XML 条目: {len(xml_entries)} (前置排布)")
    print(f"    - 媒体/素材条目: {len(media_entries)} (后置排布)")

    cd_records = []

    # 临时输出文件写入
    with open(out_pptx, 'wb') as dst:
        # Step 2: 优先写入全部 XML 结构文件（确保所有 Slide 的 Offset < 10MB，规避 32 位有符号寻址溢出）
        for info in xml_entries:
            fname = info.filename
            raw_bytes = z_orig.read(fname)

            # 如果是需要翻译替换的 Slide XML
            if fname in trans_slide_names:
                trans_raw = z_trans.read(fname)
                data = _inject_translated_text(raw_bytes, trans_raw)
            else:
                data = raw_bytes

            offset = dst.tell()
            fname_bytes = fname.encode('utf-8')
            is_compressed = (info.compress_type != 0)

            entry_bytes, csz, usz, crc, method, flags = _build_local_file_entry(
                fname_bytes, data, compress=is_compressed,
                flags=info.flag_bits,
                mod_time=info.date_time[3] << 11 | info.date_time[4] << 5 | (info.date_time[5] // 2),
                mod_date=(info.date_time[0] - 1980) << 9 | info.date_time[1] << 5 | info.date_time[2]
            )
            dst.write(entry_bytes)

            cd_entry = _build_cd_entry(
                fname_bytes, compress_method=method, flags=flags,
                mod_time=info.date_time[3] << 11 | info.date_time[4] << 5 | (info.date_time[5] // 2),
                mod_date=(info.date_time[0] - 1980) << 9 | info.date_time[1] << 5 | info.date_time[2],
                crc=crc, compressed_size=csz, uncompressed_size=usz,
                local_header_offset=offset, external_attr=info.external_attr
            )
            cd_records.append(cd_entry)

        xml_boundary = dst.tell()
        print(f"[✓] 全部核心 XML 写入完成，总体积: {xml_boundary / (1024*1024):.2f} MB (安全置于前 2GB 内)")

        # Step 3: 写入全部媒体文件（视频流不压缩，ZIP_STORED）
        for info in media_entries:
            fname = info.filename
            data = z_orig.read(fname)
            offset = dst.tell()
            fname_bytes = fname.encode('utf-8')

            # 媒体保持原有压缩模式（视频 compress_type=0 不压缩）
            is_compressed = (info.compress_type != 0)
            entry_bytes, csz, usz, crc, method, flags = _build_local_file_entry(
                fname_bytes, data, compress=is_compressed,
                flags=info.flag_bits,
                mod_time=info.date_time[3] << 11 | info.date_time[4] << 5 | (info.date_time[5] // 2),
                mod_date=(info.date_time[0] - 1980) << 9 | info.date_time[1] << 5 | info.date_time[2]
            )
            dst.write(entry_bytes)

            cd_entry = _build_cd_entry(
                fname_bytes, compress_method=method, flags=flags,
                mod_time=info.date_time[3] << 11 | info.date_time[4] << 5 | (info.date_time[5] // 2),
                mod_date=(info.date_time[0] - 1980) << 9 | info.date_time[1] << 5 | info.date_time[2],
                crc=crc, compressed_size=csz, uncompressed_size=usz,
                local_header_offset=offset, external_attr=info.external_attr
            )
            cd_records.append(cd_entry)

        # Step 4: 写入 Central Directory 与 EOCD
        cd_offset = dst.tell()
        cd_blob = b''.join(cd_records)
        dst.write(cd_blob)
        cd_size = len(cd_blob)

        eocd = _build_eocd(len(cd_records), cd_size, cd_offset)
        dst.write(eocd)

    z_orig.close()
    z_trans.close()

    new_size = os.path.getsize(out_pptx)
    orig_size = os.path.getsize(orig_pptx)
    print(f"\n[✓] 回填完成！生成文件: {out_pptx}")
    print(f"    - 原版大小: {orig_size / (1024*1024):.2f} MB")
    print(f"    - 新版大小: {new_size / (1024*1024):.2f} MB")

    # 验证
    with open(out_pptx, 'rb') as f:
        f.seek(max(0, new_size - 65536))
        tail = f.read()
    has_zip64 = b'\x50\x4b\x06\x06' in tail
    has_zip32 = b'\x50\x4b\x05\x06' in tail
    print(f"    - ZIP64 EOCD: {'有 ⚠' if has_zip64 else '无 ✓'}")
    print(f"    - ZIP32 EOCD: {'有 ✓' if has_zip32 else '无 ✗'}")

    try:
        with zipfile.ZipFile(out_pptx, 'r') as zv:
            bad = zv.testzip()
            if bad:
                print(f"[✗] 校验错误条目: {bad}")
            else:
                print(f"[✓] ZIP 校验通过: {len(zv.infolist())} 条目完好")
    except Exception as e:
        print(f"[⚠] 校验异常: {e}")


def main():
    parser = argparse.ArgumentParser(description="DeepL 翻译文本与原版大图/视频回填工具")
    parser.add_argument("--orig", required=True, help="原始未压缩 PPTX 路径")
    parser.add_argument("--trans", required=True, help="DeepL 翻译后的 PPTX 路径")
    parser.add_argument("-o", "--output", help="合并后的目标 PPTX 输出路径")

    args = parser.parse_args()
    orig = os.path.abspath(args.orig)
    trans = os.path.abspath(args.trans)

    if args.output:
        out = os.path.abspath(args.output)
    else:
        base, ext = os.path.splitext(trans)
        out = f"{base}_with_original_media{ext}"

    transplant_text(orig_pptx=orig, trans_pptx=trans, out_pptx=out)


if __name__ == "__main__":
    main()
