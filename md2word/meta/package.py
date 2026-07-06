from __future__ import annotations

import hashlib
import json
import posixpath
import re
import struct
import zipfile
from base64 import b64decode, b64encode
from io import BytesIO
from pathlib import Path, PurePosixPath
from typing import Any, Dict, Iterable, List, Optional, Tuple

from lxml import etree

from md2word.meta.semantic import build_semantic_document


META_FILENAME = "document.json"
META_FORMAT = "md2word.docx-package-meta"
META_VERSION = 1


def extract_docx_metadata(docx_path: str, output_dir: str) -> Dict[str, Any]:
    return DocxMetaExtractor().extract(docx_path, output_dir)


def restore_docx_from_metadata(meta_path: str, output_path: str) -> str:
    return DocxMetaRenderer().restore(meta_path, output_path)


def verify_docx_metadata_roundtrip(
    docx_path: str,
    output_dir: str,
    restored_name: Optional[str] = None,
) -> Dict[str, Any]:
    src = Path(docx_path)
    if not src.exists():
        raise FileNotFoundError(docx_path)

    base_dir = Path(output_dir)
    meta_dir = base_dir / "meta"
    restored_path = base_dir / (restored_name or f"{src.stem}.restored.docx")

    DocxMetaExtractor().extract(str(src), str(meta_dir))
    DocxMetaRenderer().restore(str(meta_dir), str(restored_path))

    source_bytes = src.read_bytes()
    restored_bytes = restored_path.read_bytes()
    source_sha = hashlib.sha256(source_bytes).hexdigest()
    restored_sha = hashlib.sha256(restored_bytes).hexdigest()
    return {
        "source": str(src),
        "metadata_dir": str(meta_dir),
        "restored": str(restored_path),
        "byte_identical": source_bytes == restored_bytes,
        "source_size": len(source_bytes),
        "restored_size": len(restored_bytes),
        "source_sha256": source_sha,
        "restored_sha256": restored_sha,
    }


class DocxMetaExtractor:
    """Extract a DOCX package into structured JSON plus resource files."""

    def extract(self, docx_path: str, output_dir: str) -> Dict[str, Any]:
        src = Path(docx_path)
        if not src.exists():
            raise FileNotFoundError(docx_path)

        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        source_bytes = src.read_bytes()
        entries = []
        xml_parts: Dict[str, bytes] = {}
        resource_map: Dict[str, str] = {}
        with zipfile.ZipFile(src, "r") as package:
            package_metadata = {
                "comment_base64": _bytes_to_b64(package.comment),
            }
            exact_package = _extract_exact_package(package, source_bytes, out_dir)
            if exact_package:
                package_metadata["exact"] = exact_package
            for zip_info in package.infolist():
                entry = self._entry_metadata(zip_info)
                if zip_info.is_dir():
                    entry["kind"] = "directory"
                    entries.append(entry)
                    continue

                data = package.read(zip_info)
                if _looks_like_xml_part(zip_info.filename, data):
                    entry["kind"] = "xml"
                    entry["xml"] = _xml_bytes_to_metadata(data)
                    entry["sha256"] = hashlib.sha256(data).hexdigest()
                    entry["size"] = len(data)
                    xml_parts[zip_info.filename] = data
                else:
                    entry["kind"] = "resource"
                    resource_ref = _resource_ref(zip_info.filename)
                    resource_path = _safe_join(out_dir, resource_ref)
                    resource_path.parent.mkdir(parents=True, exist_ok=True)
                    resource_path.write_bytes(data)
                    entry["resource"] = resource_ref
                    entry["sha256"] = hashlib.sha256(data).hexdigest()
                    entry["size"] = len(data)
                    resource_map[zip_info.filename] = resource_ref
                entries.append(entry)

        metadata: Dict[str, Any] = {
            "format": META_FORMAT,
            "version": META_VERSION,
            "source": {
                "filename": src.name,
                "size": len(source_bytes),
                "sha256": hashlib.sha256(source_bytes).hexdigest(),
            },
            "package": package_metadata,
            "document": build_semantic_document(xml_parts, resource_map),
            "entries": entries,
        }

        meta_file = out_dir / META_FILENAME
        with meta_file.open("w", encoding="utf-8") as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)
            f.write("\n")

        return metadata

    @staticmethod
    def _entry_metadata(zip_info: zipfile.ZipInfo) -> Dict[str, Any]:
        return {
            "path": zip_info.filename,
            "compress_type": zip_info.compress_type,
            "date_time": list(zip_info.date_time),
            "external_attr": zip_info.external_attr,
            "internal_attr": zip_info.internal_attr,
            "create_system": zip_info.create_system,
            "create_version": zip_info.create_version,
            "extract_version": zip_info.extract_version,
            "flag_bits": zip_info.flag_bits,
            "volume": zip_info.volume,
            "extra_base64": _bytes_to_b64(zip_info.extra),
            "comment_base64": _bytes_to_b64(zip_info.comment),
            "file_size": zip_info.file_size,
            "compress_size": zip_info.compress_size,
            "crc": zip_info.CRC,
        }


class DocxMetaRenderer:
    """Restore a DOCX package from structured metadata and resources."""

    def restore(self, meta_path: str, output_path: str) -> str:
        metadata, base_dir = _load_metadata(meta_path)
        _validate_metadata(metadata)

        out = Path(output_path)
        if out.parent:
            out.parent.mkdir(parents=True, exist_ok=True)

        if self._restore_exact_package(metadata, base_dir, out):
            return str(out)

        with zipfile.ZipFile(out, "w") as package:
            for entry in metadata.get("entries", []):
                zip_info = self._zip_info_from_entry(entry)
                kind = entry.get("kind")
                if kind == "directory":
                    package.writestr(zip_info, b"")
                elif kind == "xml":
                    package.writestr(zip_info, _metadata_to_xml_bytes(entry["xml"]))
                elif kind == "resource":
                    resource_path = _safe_join(base_dir, entry["resource"])
                    data = resource_path.read_bytes()
                    expected = entry.get("sha256")
                    if expected and hashlib.sha256(data).hexdigest() != expected:
                        raise ValueError(f"Resource checksum mismatch: {entry['resource']}")
                    package.writestr(zip_info, data)
                else:
                    raise ValueError(f"Unknown metadata entry kind: {kind}")
            package_info = metadata.get("package", {})
            if package_info.get("comment_base64"):
                package.comment = _b64_to_bytes(package_info["comment_base64"])

        return str(out)

    def _restore_exact_package(self, metadata: Dict[str, Any], base_dir: Path, output_path: Path) -> bool:
        exact = metadata.get("package", {}).get("exact")
        if not exact:
            return False

        entries = metadata.get("entries", [])
        exact_entries = exact.get("entries", [])
        if len(entries) != len(exact_entries):
            return False

        current_by_path = {entry.get("path"): entry for entry in entries}
        local_segments: List[bytes] = [_b64_to_bytes(exact.get("prefix_base64", ""))]
        for exact_entry in exact_entries:
            path = exact_entry.get("path")
            entry = current_by_path.get(path)
            if not entry or entry.get("kind") != exact_entry.get("kind"):
                return False
            current_data = self._entry_data(entry, base_dir)
            if len(current_data) != exact_entry.get("size"):
                return False
            expected_sha = exact_entry.get("sha256")
            if expected_sha and hashlib.sha256(current_data).hexdigest() != expected_sha:
                return False

            local_segments.append(_b64_to_bytes(exact_entry.get("local_header_base64", "")))
            payload_ref = exact_entry.get("compressed_payload")
            if payload_ref:
                payload_path = _safe_join(base_dir, payload_ref)
                if not payload_path.exists():
                    return False
                payload = payload_path.read_bytes()
                if len(payload) != exact_entry.get("compressed_size"):
                    return False
                expected_payload_sha = exact_entry.get("compressed_sha256")
                if expected_payload_sha and hashlib.sha256(payload).hexdigest() != expected_payload_sha:
                    return False
                local_segments.append(payload)
            local_segments.append(_b64_to_bytes(exact_entry.get("data_descriptor_base64", "")))

        output_path.write_bytes(b"".join(local_segments) + _b64_to_bytes(exact["central_directory_base64"]))
        expected_archive_sha = exact.get("archive_sha256")
        if expected_archive_sha and hashlib.sha256(output_path.read_bytes()).hexdigest() != expected_archive_sha:
            output_path.unlink(missing_ok=True)
            return False
        return True

    @staticmethod
    def _entry_data(entry: Dict[str, Any], base_dir: Path) -> bytes:
        kind = entry.get("kind")
        if kind == "directory":
            return b""
        if kind == "xml":
            return _metadata_to_xml_bytes(entry["xml"])
        if kind == "resource":
            resource_path = _safe_join(base_dir, entry["resource"])
            data = resource_path.read_bytes()
            expected = entry.get("sha256")
            if expected and hashlib.sha256(data).hexdigest() != expected:
                raise ValueError(f"Resource checksum mismatch: {entry['resource']}")
            return data
        raise ValueError(f"Unknown metadata entry kind: {kind}")

    @staticmethod
    def _zip_info_from_entry(entry: Dict[str, Any]) -> zipfile.ZipInfo:
        date_time = tuple(entry.get("date_time") or (1980, 1, 1, 0, 0, 0))
        zip_info = zipfile.ZipInfo(entry["path"], date_time=date_time)
        zip_info.compress_type = entry.get("compress_type", zipfile.ZIP_DEFLATED)
        zip_info.external_attr = entry.get("external_attr", 0)
        zip_info.internal_attr = entry.get("internal_attr", 0)
        zip_info.create_system = entry.get("create_system", zip_info.create_system)
        zip_info.create_version = entry.get("create_version", zip_info.create_version)
        zip_info.extract_version = entry.get("extract_version", zip_info.extract_version)
        zip_info.flag_bits = entry.get("flag_bits", zip_info.flag_bits)
        zip_info.volume = entry.get("volume", 0)
        zip_info.extra = _b64_to_bytes(entry.get("extra_base64", ""))
        zip_info.comment = _b64_to_bytes(entry.get("comment_base64", ""))
        return zip_info


def _extract_exact_package(
    package: zipfile.ZipFile,
    source_bytes: bytes,
    output_dir: Path,
) -> Dict[str, Any]:
    central_directory_offset = _central_directory_offset(source_bytes)
    if central_directory_offset is None:
        return {}

    infos = sorted(package.infolist(), key=lambda item: item.header_offset)
    exact_entries = []
    for index, zip_info in enumerate(infos):
        next_offset = (
            infos[index + 1].header_offset
            if index + 1 < len(infos)
            else central_directory_offset
        )
        exact_entry = _exact_entry_metadata(
            package,
            source_bytes,
            output_dir,
            zip_info,
            next_offset,
            index,
        )
        if not exact_entry:
            return {}
        exact_entries.append(exact_entry)

    first_offset = infos[0].header_offset if infos else central_directory_offset
    return {
        "archive_sha256": hashlib.sha256(source_bytes).hexdigest(),
        "prefix_base64": _bytes_to_b64(source_bytes[:first_offset]),
        "central_directory_base64": _bytes_to_b64(source_bytes[central_directory_offset:]),
        "entries": exact_entries,
    }


def _exact_entry_metadata(
    package: zipfile.ZipFile,
    source_bytes: bytes,
    output_dir: Path,
    zip_info: zipfile.ZipInfo,
    next_offset: int,
    index: int,
) -> Dict[str, Any]:
    header_offset = zip_info.header_offset
    local_fixed_header = source_bytes[header_offset:header_offset + 30]
    if len(local_fixed_header) != 30:
        return {}
    try:
        (
            signature,
            _version,
            _flag_bits,
            _compress_type,
            _mod_time,
            _mod_date,
            _crc,
            _compressed_size,
            _file_size,
            filename_length,
            extra_length,
        ) = struct.unpack("<IHHHHHIIIHH", local_fixed_header)
    except struct.error:
        return {}
    if signature != 0x04034B50:
        return {}

    payload_start = header_offset + 30 + filename_length + extra_length
    compressed_end = payload_start + zip_info.compress_size
    if payload_start > len(source_bytes) or compressed_end > next_offset:
        return {}

    local_header = source_bytes[header_offset:payload_start]
    compressed_payload = source_bytes[payload_start:compressed_end]
    data_descriptor = source_bytes[compressed_end:next_offset]
    if zip_info.is_dir():
        data = b""
        kind = "directory"
    else:
        data = package.read(zip_info)
        kind = "xml" if _looks_like_xml_part(zip_info.filename, data) else "resource"

    compressed_sha = hashlib.sha256(compressed_payload).hexdigest() if compressed_payload else ""
    payload_ref = ""
    if compressed_payload:
        payload_ref = posixpath.join(
            "resources",
            ".zip-payloads",
            f"{index:04d}-{compressed_sha}.bin",
        )
        payload_path = _safe_join(output_dir, payload_ref)
        payload_path.parent.mkdir(parents=True, exist_ok=True)
        payload_path.write_bytes(compressed_payload)

    return {
        "path": zip_info.filename,
        "kind": kind,
        "size": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "compressed_size": len(compressed_payload),
        "compressed_sha256": compressed_sha,
        "compressed_payload": payload_ref,
        "local_header_base64": _bytes_to_b64(local_header),
        "data_descriptor_base64": _bytes_to_b64(data_descriptor),
    }


def _central_directory_offset(data: bytes) -> Optional[int]:
    search_start = max(0, len(data) - (65535 + 22))
    eocd_offset = data.rfind(b"PK\x05\x06", search_start)
    if eocd_offset < 0 or eocd_offset + 22 > len(data):
        return None
    try:
        (
            _signature,
            _disk_number,
            _central_directory_disk,
            _disk_entries,
            _total_entries,
            _central_directory_size,
            central_directory_offset,
            comment_length,
        ) = struct.unpack_from("<IHHHHIIH", data, eocd_offset)
    except struct.error:
        return None
    if eocd_offset + 22 + comment_length != len(data):
        return None
    if central_directory_offset == 0xFFFFFFFF:
        return None
    if central_directory_offset > len(data):
        return None
    return central_directory_offset


def _load_metadata(path: str) -> Tuple[Dict[str, Any], Path]:
    p = Path(path)
    if p.is_dir():
        meta_file = p / META_FILENAME
        base_dir = p
    else:
        meta_file = p
        base_dir = p.parent

    with meta_file.open("r", encoding="utf-8") as f:
        return json.load(f), base_dir


def _validate_metadata(metadata: Dict[str, Any]) -> None:
    if metadata.get("format") != META_FORMAT:
        raise ValueError("Unsupported metadata format")
    if metadata.get("version") != META_VERSION:
        raise ValueError("Unsupported metadata version")
    if not isinstance(metadata.get("entries"), list):
        raise ValueError("Metadata is missing entries")


def _looks_like_xml_part(name: str, data: bytes) -> bool:
    lower = name.lower()
    if lower.endswith(".xml") or lower.endswith(".rels"):
        return True
    return data.lstrip().startswith(b"<?xml")


def _xml_bytes_to_metadata(data: bytes) -> Dict[str, Any]:
    parser = etree.XMLParser(
        resolve_entities=False,
        remove_blank_text=False,
        remove_comments=False,
        strip_cdata=False,
        huge_tree=True,
    )
    tree = etree.parse(BytesIO(data), parser)
    docinfo = tree.docinfo
    metadata = {
        "declaration": {
            "present": data.lstrip().startswith(b"<?xml"),
            "version": docinfo.xml_version or "1.0",
            "encoding": docinfo.encoding or "UTF-8",
            "standalone": _declared_standalone(data),
        },
        "trailing_newline": data.endswith(b"\n"),
        "root": _node_to_metadata(tree.getroot(), {}),
    }
    preamble = _xml_preamble(data)
    if preamble:
        metadata["preamble"] = preamble.decode("ascii")
    return metadata


def _metadata_to_xml_bytes(xml_metadata: Dict[str, Any]) -> bytes:
    root = _metadata_to_node(xml_metadata["root"])
    declaration = xml_metadata.get("declaration", {})
    preamble = xml_metadata.get("preamble")
    if preamble is not None:
        data = preamble.encode("ascii") + etree.tostring(
            root,
            encoding=declaration.get("encoding") or "UTF-8",
            xml_declaration=False,
        )
    else:
        kwargs: Dict[str, Any] = {
            "encoding": declaration.get("encoding") or "UTF-8",
            "xml_declaration": bool(declaration.get("present", True)),
        }
        if declaration.get("standalone") is not None:
            kwargs["standalone"] = bool(declaration["standalone"])
        data = etree.tostring(root, **kwargs)
    if xml_metadata.get("trailing_newline") and not data.endswith(b"\n"):
        data += b"\n"
    return data


def _node_to_metadata(node: etree._Element, inherited_ns: Dict[Optional[str], str]) -> Dict[str, Any]:
    if isinstance(node, etree._Comment):
        return {"kind": "comment", "text": node.text, "tail": node.tail}
    if isinstance(node, etree._ProcessingInstruction):
        return {
            "kind": "processing_instruction",
            "target": node.target,
            "text": node.text,
            "tail": node.tail,
        }

    current_ns = dict(node.nsmap or {})
    declared_ns = {
        _json_ns_key(prefix): uri
        for prefix, uri in current_ns.items()
        if inherited_ns.get(prefix) != uri
    }
    item: Dict[str, Any] = {
        "kind": "element",
        "tag": _name_to_metadata(node.tag, current_ns),
    }
    if declared_ns:
        item["namespaces"] = declared_ns
    if node.attrib:
        item["attributes"] = [
            {
                "name": _name_to_metadata(name, current_ns),
                "value": value,
            }
            for name, value in node.attrib.items()
        ]
    if node.text is not None:
        item["text"] = node.text
    children = [_node_to_metadata(child, current_ns) for child in node]
    if children:
        item["children"] = children
    if node.tail is not None:
        item["tail"] = node.tail
    return item


def _metadata_to_node(item: Dict[str, Any]) -> etree._Element:
    kind = item.get("kind")
    if kind == "comment":
        node = etree.Comment(item.get("text") or "")
    elif kind == "processing_instruction":
        node = etree.ProcessingInstruction(item["target"], item.get("text") or "")
    elif kind == "element":
        nsmap = _json_to_nsmap(item.get("namespaces", {}))
        node = etree.Element(_metadata_to_qname(item["tag"]), nsmap=nsmap or None)
        for attr in item.get("attributes", []):
            node.set(_metadata_to_qname(attr["name"]), attr["value"])
        if "text" in item:
            node.text = item["text"]
        for child in item.get("children", []):
            node.append(_metadata_to_node(child))
    else:
        raise ValueError(f"Unknown XML node kind: {kind}")

    if "tail" in item:
        node.tail = item["tail"]
    return node


def _name_to_metadata(qname: str, nsmap: Dict[Optional[str], str]) -> Dict[str, Optional[str]]:
    namespace = None
    local_name = qname
    if qname.startswith("{"):
        namespace, local_name = qname[1:].split("}", 1)
    prefix = _prefix_for_namespace(nsmap, namespace)
    if namespace and prefix:
        display = f"{prefix}:{local_name}"
    elif namespace:
        display = f"{{{namespace}}}{local_name}"
    else:
        display = local_name
    return {
        "display": display,
        "namespace": namespace,
        "name": local_name,
        "prefix": prefix,
    }


def _metadata_to_qname(name: Dict[str, Optional[str]]) -> str:
    namespace = name.get("namespace")
    local_name = name["name"]
    if namespace:
        return f"{{{namespace}}}{local_name}"
    return local_name


def _prefix_for_namespace(nsmap: Dict[Optional[str], str], namespace: Optional[str]) -> Optional[str]:
    if not namespace:
        return None
    for prefix, uri in nsmap.items():
        if uri == namespace:
            return prefix or ""
    return None


def _declared_standalone(data: bytes) -> Optional[bool]:
    match = re.match(br"\s*<\?xml[^>]*standalone=[\"'](yes|no)[\"']", data)
    if not match:
        return None
    return match.group(1) == b"yes"


def _xml_preamble(data: bytes) -> bytes:
    match = re.match(br"\A(\s*<\?xml.*?\?>)([ \t\r\n]*)", data, re.DOTALL)
    if match:
        return match.group(0)
    match = re.match(br"\A([ \t\r\n]+)(?=<)", data)
    return match.group(1) if match else b""


def _json_ns_key(prefix: Optional[str]) -> str:
    return "" if prefix is None else prefix


def _json_to_nsmap(value: Dict[str, str]) -> Dict[Optional[str], str]:
    return {(None if prefix == "" else prefix): uri for prefix, uri in value.items()}


def _resource_ref(package_path: str) -> str:
    _assert_safe_posix_path(package_path)
    return posixpath.join("resources", package_path)


def _safe_join(root: Path, relative: str) -> Path:
    parts = _safe_parts(relative)
    return root.joinpath(*parts)


def _safe_parts(relative: str) -> Iterable[str]:
    _assert_safe_posix_path(relative)
    return PurePosixPath(relative).parts


def _assert_safe_posix_path(relative: str) -> None:
    pure = PurePosixPath(relative)
    if pure.is_absolute() or ".." in pure.parts:
        raise ValueError(f"Unsafe package path: {relative}")


def _bytes_to_b64(value: bytes) -> str:
    return b64encode(value).decode("ascii") if value else ""


def _b64_to_bytes(value: str) -> bytes:
    return b64decode(value.encode("ascii")) if value else b""
