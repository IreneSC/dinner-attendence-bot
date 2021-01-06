from discord.ext.commands import Bot, has_permissions
from discord import File
from time import localtime, strftime
from xlsxwriter import Workbook
import os
import io

bot = Bot(command_prefix='!')

house_attendees = {}
logging = False


def log_channel_change(member_id, channel, log_string):
    house = str(channel.category)
    table = str(channel)
    if house not in house_attendees:
        house_attendees[house] = {}
    if table not in house_attendees[house]:
        house_attendees[house][table] = {}
    if member_id not in house_attendees[house][table]:
        house_attendees[house][table][member_id] = ""
    house_attendees[house][table][member_id] += " / " + log_string


@bot.command(name='startdinner')
@has_permissions(administrator=True)
async def start_logging(context):
    global house_attendees
    global logging
    house_attendees = {}
    logging = True
    await context.send('Started Logging')


@bot.command(name='enddinner')
@has_permissions(administrator=True)
async def stop_logging(context):
    await context.send('Stopped Logging')
    global logging
    logging = False
    
    output = io.BytesIO()
    wb = Workbook(output, {'in_memory': True})
    
    global house_attendees
    for house in house_attendees:
        frosh = set()
        table_attendance = {}
        for table in house_attendees[house]:
            table_attendance[table] = []
            for member_id in house_attendees[house][table]:
                try:
                    member = await context.guild.fetch_member(member_id)
                    print("Fetched Member", member.display_name)
                    table_attendance[table].append(member.display_name + house_attendees[house][table][member.id])
                    if "Frosh" in [r.name for r in member.roles]:
                        frosh.add(member.display_name + "(" + str(member) + ")")
                except:
                    print("That frosh either changed roles or left the server.")
                    continue
        worksheet = wb.add_worksheet(house)
        worksheet.write_column(0, 0, ["All Frosh"] + list(frosh))
        for i,table in enumerate(table_attendance.keys()):
            worksheet.write_column(0, i + 1, [table] + table_attendance[table])
    wb.close()

    output.seek(0)
    await context.send(content="Sending you the Excel workbook now!", 
                       file=File(fp=output, 
                                 filename="Dinner Attendance " + strftime("%m-%d %H:%M",localtime()) + ".xlsx"))


@bot.event
async def on_voice_state_update(member, before, after):
    global logging
    if not logging:
        return
    if before.channel == after.channel:
        return
    if before.channel is not None:
        log_channel_change(member.id, before.channel, "Left At " + strftime("%H:%M", localtime()))
    if after.channel is not None:
        log_channel_change(member.id, after.channel, "Joined At " + strftime("%H:%M", localtime()))


bot.run(os.getenv('DISCORD_TOKEN'))
