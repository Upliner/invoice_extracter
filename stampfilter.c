#include <stdio.h>
#include <stdlib.h>
#include <string.h>

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
     return (blue*2-green/2-red/2 > 256) ? 1 : 0;
}
int main(int argc, char* argv[]) {
    int w, h, i, x, y, fpos;
    int start, end; // Starts and ends of edited area
    unsigned char *prevrow, *currow, *nextrow; // Rows of pixels
    unsigned char *cur, *prevcur; // Current pixel
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
    prevrow = malloc(w);
    currow  = malloc(w*3);
    nextrow = malloc(w*3);
    if (!prevrow || !currow || !nextrow) {
        fprintf(stderr, "Cannot allocate memory\n");
        return 1;
    }
    memset(prevrow, 0, w);
    fread(currow, 3, w, f);
    start = end = -1;
    fpos = 0;
    for (y = 1; y < h; y++) {
        start = end = -1;
        i = fread(nextrow, 3, w, f);
        if (i < w) {
            fprintf(stderr, "Warning: unexpected end of file\n");
            return 0;
        }
        for (x = 0, cur = currow, prevcur = prevrow; x < w; x++, prevcur++, cur += 3)
            if (shouldFilter(cur)) {
                /* If we should filter current pixel, check 8 neighbor pixels.
                   Apply filter only if at least 4 of them are blue too */
                int numpixels = 0;
                numpixels += prevcur[-1] + prevcur[0] + prevcur[1]; // Previous row
                numpixels += (end == (x-1)) ? 1 : shouldFilter(cur-3) + shouldFilter(cur+3); // Current row
                if (numpixels < 4) // Next row
                   numpixels += shouldFilter(&nextrow[(x-1)*3]) + shouldFilter(&nextrow[x*3]) + shouldFilter(&nextrow[(x+1)*3]);
                if (numpixels >= 4) {
                    if (start == -1) start = x;
                    end = x;
                    cur[0] = cur[1] = cur[2] = 255;
                }
                *prevcur = 1;
            } else *prevcur = 0;
        if (start != -1) {
            fpos = ftell(f);
            if (fpos == -1) break;
            if (fseek(f, (start - w*2)*3, SEEK_CUR)) { fpos = -1; break; }
            fwrite(currow + start*3, 3, end - start, f);
            fseek(f, fpos, SEEK_SET);
        }
        cur = currow;
        currow = nextrow;
        nextrow = cur;
    }
    if (fpos == -1)
        fprintf(stderr, "File seems to be unseekable\n");
    free(prevrow);
    free(currow);
    free(nextrow);
    return fpos == -1 ? 1 : 0;
}
