import requests
import telebot

import abcxyz_method
import credits
import copy
import io


bot = telebot.TeleBot(credits.TOKEN)


# image to send for help
with open("imgs/example.png", 'rb') as file:
    example_img = io.BytesIO(file.read())

# example file to send for help
with open("example_file.xlsx", 'rb') as file:
    example_file = io.BytesIO(file.read())
    example_file.name = "example_file.xlsx"


@bot.message_handler(content_types=["document"])
def main(msg: telebot.types.Message):

    try:
        # getting document from message
        document = requests.get(f"https://api.telegram.org/file/bot{credits.TOKEN}/{bot.get_file(msg.document.file_id).file_path}").content
        
        bot.reply_to(msg, "processing...")

        # make calculations, return result of two files
        result = abcxyz_method.abcmethod(document)

        # send result documents
        for i in result:
            new_file = telebot.types.InputFile(i)
            bot.send_document(chat_id=msg.chat.id, document=new_file, reply_to_message_id=msg.message_id)

    except:
        bot.reply_to(msg, "error processing file, check file format")


# send help message and example file
@bot.message_handler(commands=["help", "start"])
def help(msg: telebot.types.Message):

    bot.send_photo(msg.chat.id, 
                    copy.copy(example_img), 
                    caption="this is bot to analyse xlsx file and do economical abcxyz analyse method, file should be formatted as on image and contain from 1 to 16 months\n\njust send me a .xlsx file"
                    )
    
    bot.send_document(
        chat_id=msg.chat.id, 
        document=copy.copy(example_file), 
        caption="here is an example file"
    )


if __name__ == '__main__':
    bot.infinity_polling()

