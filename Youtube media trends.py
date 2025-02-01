import os
import requests
import pandas as pd
from youtubesearchpython import VideosSearch
import yt_dlp
import time
from PIL import Image, UnidentifiedImageError
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime
import streamlit as st

# Function to format duration
def format_duration(seconds):
    if isinstance(seconds, int):
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        seconds = seconds % 60
        return f"{hours:02}:{minutes:02}:{seconds:02}"
    return "N/A"

# Function to download images
def download_image(url, save_path):
    try:
        response = requests.get(url, stream=True, timeout=10)
        if response.status_code == 200:
            with open(save_path, 'wb') as file:
                for chunk in response.iter_content(1024):
                    file.write(chunk)
            with Image.open(save_path) as img:
                img.verify()
                img.convert("RGB").save(save_path, "JPEG")  # Convert to JPEG if needed
                return save_path
    except (UnidentifiedImageError, ValueError, Exception) as e:
        print(f"Error downloading or validating image: {e}")
    return None

# Function to get video details
def get_video_details(video_url, idx):
    ydl_opts = {'quiet': True, 'force_generic_extractor': True}
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info_dict = ydl.extract_info(video_url, download=False)
        thumbnail_url = info_dict.get('thumbnail', "N/A")
        thumbnail_path = f"thumbnail_{idx}.jpg" if thumbnail_url != "N/A" else "N/A"
        if thumbnail_url != "N/A":
            download_image(thumbnail_url, thumbnail_path)

        # Format the upload date to dd/mm/yyyy, hh:mm:ss
        upload_date = info_dict.get('upload_date', "N/A")
        if upload_date != "N/A":
            try:
                upload_date = datetime.strptime(upload_date, "%Y%m%d").strftime("%d/%m/%Y, %H:%M:%S")
            except Exception as e:
                print(f"Error formatting date: {e}")
                upload_date = "N/A"

        return {
            "title": info_dict.get('title', "N/A"),
            "url": video_url,
            "channel_name": info_dict.get('uploader', "N/A"),
            "views": info_dict.get('view_count', "N/A"),
            "duration": format_duration(info_dict.get('duration', "N/A")),
            "likes": info_dict.get('like_count', "N/A"),
            "comments": info_dict.get('comment_count', "N/A"),
            "thumbnail": thumbnail_path,
            "date": upload_date,  # Updated date format
            "upload_timestamp": info_dict.get('upload_date', "N/A"),  # For sorting
        }

# Function to search YouTube
def search_youtube(keywords, max_results=20):
    video_details_list = []
    combined_results = []

    # Fetch results for all keywords and combine them
    for keyword in keywords:
        videos_search = VideosSearch(keyword, limit=max_results)
        combined_results.extend(videos_search.result()['result'])

    # Deduplicate results by video URL
    unique_results = []
    seen_urls = set()
    for video in combined_results:
        video_url = f"https://www.youtube.com/watch?v={video['id']}"
        if video_url not in seen_urls:
            seen_urls.add(video_url)
            unique_results.append(video)

    # Get details for all unique results
    for idx, video in enumerate(unique_results, start=1):
        video_url = f"https://www.youtube.com/watch?v={video['id']}"
        video_details = get_video_details(video_url, idx)
        video_details_list.append(video_details)
        time.sleep(1)  # Avoid rate limiting

    # Sort results by upload date (most recent first)
    video_details_list.sort(key=lambda x: x["upload_timestamp"], reverse=True)

    # Return only the top 20 most recent results
    return video_details_list[:max_results]

# Function to save data to Excel
def save_to_excel(data, output_excel="youtube_data.xlsx"):
    df = pd.DataFrame(data)
    # Drop the temporary 'upload_timestamp' column before saving to Excel
    df.drop(columns=["upload_timestamp"], inplace=True)
    df.to_excel(output_excel, index=False)
    wb = openpyxl.load_workbook(output_excel)
    ws = wb.active
    for idx, row in enumerate(data, start=2):
        if row["thumbnail"] != "N/A" and os.path.exists(row["thumbnail"]):
            img = XLImage(row["thumbnail"])
            img.width, img.height = 100, 50 
            ws.add_image(img, f"B{idx}")
    wb.save(output_excel)
    print(f"Excel file saved as {output_excel}")

# Streamlit App
def main():
    st.title("YouTube Media Trends in Andhra Pradesh")
    
    # Input box for keywords/hashtags
    keywords_input = st.text_input("Enter keywords/hashtags (comma-separated):")
    
    # Button to trigger data fetching
    if st.button("Get Data"):
        if keywords_input:
            # Split keywords by comma and strip whitespace
            keywords = [keyword.strip() for keyword in keywords_input.split(",")]
            
            # Fetch data
            st.write("Fetching data... Please wait.")
            data = search_youtube(keywords, max_results=20)
            
            # Display data in a table
            if data:
                df = pd.DataFrame(data)
                df.drop(columns=["upload_timestamp"], inplace=True)  # Remove sorting column
                st.write("Top 20 Recent Results:")
                st.dataframe(df)

                # Save data to Excel
                excel_file = "youtube_data.xlsx"
                save_to_excel(data, excel_file)

                # Provide a download button for the Excel file
                with open(excel_file, "rb") as file:
                    st.download_button(
                        label="Download Excel",
                        data=file,
                        file_name=excel_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.write("No results found.")
        else:
            st.write("Please enter at least one keyword or hashtag.")

# Run the Streamlit app
if __name__ == "__main__":
    main()