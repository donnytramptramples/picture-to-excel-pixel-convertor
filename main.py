from PIL import Image
import xlsxwriter
from tqdm import tqdm

def toHexa(val, channel):
    """
    Convert decimal value to hexadecimal color code where
    only channel given has a value and the other two are set
    to 00.

    val: decimal value
    channel: color channel
    """

    if channel == 'r':  
        return '#' + '{:02x}'.format(val) + '0000'
    elif channel == 'g':
        return '#00' + '{:02x}'.format(val) + '00'
    else:
        return '#0000' + '{:02x}'.format(val)

def main():
    # Ask for input image file name
    input_image = input("Enter the input image file name: ")

    # Load image
    im = Image.open(input_image)
    imsize = im.size
    print(f'Input image size: ({imsize[0]}, {imsize[1]})')

    if imsize[0] >500 or imsize[1] > 500:
        print('Downsampling image.')
        im.thumbnail((500, 500), Image.LANCZOS) # downsampling, aspect ratio stays the same
        imsize = im.size
        print(f'Image size after downsampling: ({imsize[0]}, {imsize[1]})')
    pix = im.load()

    # Create Excel workbook/worksheet
    workbook = xlsxwriter.Workbook('./output.xlsx')
    worksheet = workbook.add_worksheet('image')

    # Color excel cells
    total_pixels = imsize[0] * imsize[1] * 3
    progress_bar = tqdm(total=total_pixels, desc="Progress", unit="pixels")

    for x in range(imsize[0]):
        for y in range(0, imsize[1]*3, 3):
            colors = pix[x, y//3]
            for i, channel in enumerate(['r', 'g', 'b']):
                hexcode = toHexa(colors[i], channel=channel)
                wbformat = workbook.add_format({'bg_color': hexcode})
                worksheet.write(y+i, x, colors[i], wbformat)
                progress_bar.update(1)

    progress_bar.close()
    workbook.close()
    print('Spreadsheet complete. Saved to ./output.xlsx')

if __name__ == '__main__':
    main()
