# Week 9 · Day 3 (Wed) · Theme and Legacy: Course Mojo Lesson 6 + Closing the Play

Prather Week 2, Day 3, taught as Week 9.

| File | What it is |
|---|---|
| `Week9_Day3_Student_Packet.pdf` | 3-page universal student packet (US Letter, print single-sided) |
| `Week9_Day3_Teacher_Copy.pdf` | 13-page landscape teach-through: lesson at a glance, text + key, then one page per segment |
| `Week9_Day3_Slides.pptx` | 12 slides with speaker notes |
| `src/` | Editable sources: `student_packet.html` + `packet.css`, `teacher_copy.html` + `teacher.css`, `build_slides.py`, `build.py` |

## Rebuild after editing

```
cd src
python3 build.py
```

This rebuilds all three files. It needs `weasyprint`, `python-pptx`, `Pillow`, `poppler-utils` (`pdftoppm`) and the
TeX Gyre Adventor font (`apt-get install fonts-texgyre`). The teacher copy's packet thumbnails, with the active box
outlined in red, are regenerated from the packet each time, so packet edits flow into the teacher copy automatically.
Teacher-copy thumbnails are placed with `{{thumb:<region>}}`, where `<region>` is a `data-hl` id in the packet HTML.

## Placeholders to fill

- **Slide 2:** the link to “Anne Frank: the only existing film images.” The Prather plan names the clip but gives no URL.
