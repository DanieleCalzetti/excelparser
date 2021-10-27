import string

class read_document:

    def __init__(self, sheet):
        self.sheet = sheet


    def check_row_used(self, first_post):
        self.row_in_use = []
        count = int(first_post)
        row = self.sheet["A" + str(count)].value
        while row != None: 
            self.row_in_use.append(str(count))
            count += 1
            row = self.sheet["A" + str(count)].value
        self.row_in_use

    def find_header(self, coverage_header, like_header, row_header, impression_header, hashtag_header):
        columns = list(string.ascii_uppercase)
        self.column_coverage = ""
        self.column_hashtag = ""
        self.column_like = ""
        self.column_hashtag = ""
        for letter in columns:
            header = self.sheet[letter + str(row_header)].value
            if header == coverage_header:
                self.column_coverage = letter
            elif header == like_header:
                self.column_like = letter
            elif header == impression_header:
                self.column_impression = letter
            elif header == hashtag_header:
                self.column_hashtag = letter
        return self.column_impression, self.column_coverage, self.column_hashtag, self.column_like
            
    def find_top_post(self):
        total_coverage = {}
        total_like = {}
        total_impression = {}       
        for num in self.row_in_use:
            coverage_value = self.sheet[self.column_coverage + str(num)].value
            total_coverage[num] = coverage_value
            total_coverage_filtered = dict(filter(lambda item: item[1] is not None, total_coverage.items()))
            self.max_coverage = max(total_coverage_filtered.values())
            self.row_max_coverage = max(total_coverage_filtered, key=total_coverage_filtered.get)

            like_value = self.sheet[self.column_like + str(num)].value
            total_like[num] = like_value
            total_like_filtered = dict(filter(lambda item: item[1] is not None, total_like.items()))
            self.max_like = max(total_like_filtered.values())
            self.row_max_like = max(total_like_filtered, key=total_like_filtered.get)       

            impression_value = self.sheet[self.column_impression + str(num)].value
            total_impression[num] = impression_value
            total_impression_filtered = dict(filter(lambda item: item[1] is not None, total_impression.items()))
            self.max_impression = max(total_impression_filtered.values())
            self.row_max_impression = max(total_impression_filtered, key=total_impression_filtered.get)
        return self.row_max_coverage, self.row_max_impression, self.row_max_like

      
    def compare_hashtag(self):
        hashtags_max_coverage = list(self.sheet[self.column_hashtag + str(self.row_max_coverage)].value.split())
        hashtags_max_like = list(self.sheet[self.column_hashtag + str(self.row_max_like)].value.split())
        hashtags_max_impression = list(self.sheet[self.column_hashtag + str(self.row_max_impression)].value.split())
        self.common_hashtags = []
        for hashtag in hashtags_max_coverage:
            if hashtag in hashtags_max_like and hashtag not in self.common_hashtags:
                self.common_hashtags.append(hashtag)
            if hashtag in hashtags_max_impression and hashtag not in self.common_hashtags:
                self.common_hashtags.append(hashtag)
        for hashtag in hashtags_max_impression:
            if hashtag in hashtags_max_like and hashtag not in self.common_hashtags:
                self.common_hashtags.append(hashtag)
        self.str_common_hashtags = ' '.join([str(elem) for elem in self.common_hashtags])
        if self.str_common_hashtags == "":
            self.str_common_hashtags = "There aren't common hashtags"
        return self.str_common_hashtags
