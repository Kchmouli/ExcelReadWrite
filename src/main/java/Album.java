
import java.util.Date;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author KantipudiChandraMouli
 */
public class Album implements Comparable{

    private String albumName;
    private String genre;
    private String artist;
    private Date releasedate;
    private Integer criticScore;

    public Album(String albumName, String genre, String artist, Date releasedate, int criticScore) {
        this.albumName = albumName;
        this.genre = genre;
        this.artist = artist;
        this.releasedate = releasedate;
        this.criticScore = criticScore;
    }

    public String getAlbumName() {
        return albumName;
    }

    public String getGenre() {
        return genre;
    }

    public String getArtist() {
        return artist;
    }

    public Date getReleasedate() {
        return releasedate;
    }

    public int getCriticScore() {
        return criticScore;
    }
      
    
    @Override
    public int compareTo(Object o) {
        if(o instanceof Album){
            Album other = (Album)o;
            int genreComp = this.getGenre().compareTo(other.getGenre());
            if(genreComp == 0)
                return other.criticScore.compareTo(this.criticScore);
            return genreComp;
        }
        return -1;
    }
    
    @Override
    public String toString(){
        return albumName+ ", " + genre + ", " + criticScore;
    }
}