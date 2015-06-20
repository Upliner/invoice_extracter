#include <stdio.h>
#include <stdlib.h>
#define buflen 256
#define btresh 80
#define min(a,b) ((a)<(b)?(a):(b))
char buf[buflen], *s;
FILE *f;
char *filename;
void readline() {
    do {
        s = fgets(buf, buflen, f);
        if (!s) {
            fprintf(stderr, "Bad image format in %s\n", filename);
            exit(1);
        }
        if (*buf == '#') continue; // skip comments
        return;
    } while (1);
}
int shouldFilter(unsigned char* cur) {
     if (cur[2]<btresh) return 0;
     int red = cur[0];
     int green = cur[1];
     int blue = cur[2];
     return blue*2-green/2-red/2 > 260;
}
int main(int argc, char* argv[]) {
    int w, h, i, x, y, start, end, fpos;
    unsigned char *row, *cur;
    if (argc < 2) {
        fprintf(stderr, "Usage: stampfilter image.ppm\n");
        return 1;
    }
    filename = argv[1];
    f = fopen(filename, "r+b");
    if (!f) {
        fprintf(stderr, "Usage: can't open file %s\n", filename);
        return 1;
    }
    readline();
    if (strncmp(buf, "P6\n", 3) != 0) {
        fprintf(stderr, "Bad image format: only RGB ppm images are supported\n");
        return 1;
    }
    readline();
    i = sscanf(buf, "%u %u", &w, &h);
    if (i == 0) {
        fprintf(stderr, "Can't read image dimensions\n");
        return 1;
    }
    if (i == 1) {
        readline();
        i = sscanf(buf, "%u", &h);
        if (i == 0) {
            fprintf(stderr, "Can't read image dimensions\n");
            return 1;
        }
    }
    if (fscanf(f, "%u", &i) == 0) {
        fprintf(stderr, "Can't read image color depth\n");
        return 1;
    }
    if (i != 255) {
        fprintf(stderr, "Bad color depth, must be 255\n");
        return 1;
    }
    if (fgetc(f) != 10) {
        fprintf(stderr, "Expected newline after color depth\n");
        return 1;
    }
    row = malloc(w*3);
    if (!row) {
        fprintf(stderr, "Cannot allocate memory\n");
        return 1;
    }
    for (y = 0; y < h; y++) {
        start = end = -1;
        i = fread(row, 3, w, f);
        if (i < w) {
            fprintf(stderr, "Warning: unexpected end of file\n");
            return 0;
        }
        for (x = 0, cur = row; x < w; x++, cur += 3)
            if (shouldFilter(cur))
                if (end == (x-1) || shouldFilter(cur+3)) {
                    if (start == -1) start = x;
                    end = x;
                    cur[0] = cur[1] = cur[2] = 255;
                }
        if (start != -1) {
            fpos = ftell(f);
            if (fpos == -1) goto seekerr;
            if (fseek(f, (start - w)*3, SEEK_CUR)) goto seekerr;
            fwrite(row + start*3, 3, end - start, f);
            fseek(f, fpos, SEEK_SET);
        }
    }
    free(row);
    return 0;
    seekerr:
    fprintf(stderr, "file seems to be unseekable\n");
    free(row);
    return 1;
}
