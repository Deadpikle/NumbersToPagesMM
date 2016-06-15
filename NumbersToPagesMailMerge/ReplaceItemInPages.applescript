on replace_tag(thisTag, thisString)
    tell application "Pages"
        activate
        tell the front document
            set (first placeholder text whose tag is thisTag) to thisString
        end tell
    end
end show_message